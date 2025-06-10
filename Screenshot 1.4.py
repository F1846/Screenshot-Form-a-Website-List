import os
os.environ["TK_SILENCE_DEPRECATION"] = "1"

import time
import csv
import threading
import queue  
from urllib.parse import urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
from tkinter import Toplevel  

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import openpyxl  



def init_driver(headless=True):
    chrome_options = Options()
    if headless:
        chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--window-size=1920,1080')
    chrome_options.add_argument("--blink-settings=imagesEnabled=false")
    chrome_options.add_argument("--disable-extensions")
    service = Service()
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def create_driver_pool(num, headless=True):
    pool = queue.Queue()
    for _ in range(num):
        driver = init_driver(headless)
        pool.put(driver)
    return pool



def take_screenshot(url, output_folder, stop_event, headless=True):
    """Attempts to take a screenshot using a pooled Chrome driver.
       Returns a tuple: (url, success, message, server_response).
       
       server_response will be "Success" if page loads normally,
       "Blocked or Captcha" if page content signals blocking,
       or "Error" if an exception occurred.
       
       For Instagram profiles, the function waits until the profile header image is loaded.
    """
    if stop_event.is_set():
        return (url, False, "Stopped", "N/A")
    
    url = url.strip()
    if not (url.startswith("http://") or url.startswith("https://")):
        url = "https://" + url  
    parsed_url = urlparse(url)
    domain = parsed_url.netloc.replace("www.", "")
    
    
    username = parsed_url.path.strip("/").split("/")[0] or "profile"
    filename = f"{domain}_{username}.png"
    screenshot_path = os.path.join(output_folder, filename)

    if not os.path.exists(output_folder):
        try:
            os.makedirs(output_folder)
        except Exception as e:
            return (url, False, f"Folder error for {url}: {e}", "N/A")

    try:
        driver = DRIVER_POOL.get(timeout=10)
    except Exception as e:
        return (url, False, f"Driver pool error: {e}", "N/A")

    try:
        driver.get(url)
        
        server_response = "Success"
        
        # If URL is Instagram/FB/TikTok, wait for profile header picture to load(to implemenet-captcha block)
        if "instagram.com""tiktok.com""facebook.com" in domain:
            try:
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.XPATH, '//header//img[contains(@alt, "profile picture")]'))
                )
                
                time.sleep(20)
            except Exception as e:
                server_response = "Blocked or Captcha"
        else:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, 'body'))
            )
            time.sleep(2)
            
            body_text = driver.find_element(By.TAG_NAME, "body").text.lower()
            if any(word in body_text for word in ["access denied", "forbidden", "error 403", "blocked", "verify you are human"]):
                server_response = "Blocked or Captcha"

        if driver.save_screenshot(screenshot_path):
            return (url, True, f"Screenshot saved: {screenshot_path}", server_response)
        else:
            return (url, False, f"Screenshot not saved for {url}", server_response)
    except Exception as e:
        return (url, False, f"Error processing {url}: {e}", "Url not reachable")
    finally:
        DRIVER_POOL.put(driver)



def save_results_to_csv(url_results, output_folder):
    """Saves screenshot results to a CSV file with three columns: URL, Success, and Server Response."""
    csv_path = os.path.join(output_folder, "screenshot_results.csv")
    try:
        with open(csv_path, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(['URL', 'Success', 'Server Response'])
            for result in url_results:
                writer.writerow(result)  
        print(f"[INFO] Results saved to CSV: {csv_path}")
    except Exception as e:
        print(f"[ERROR] Could not save CSV file: {e}")

def save_results_to_excel(url_results, output_folder):
    """Saves screenshot results to an Excel file with three columns: URL, Success, and Server Response."""
    excel_path = os.path.join(output_folder, "screenshot_results.xlsx")
    try:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["URL", "Success", "Server Response"])
        for result in url_results:
            ws.append(result)
        wb.save(excel_path)
        print(f"[INFO] Results saved to Excel: {excel_path}")
    except Exception as e:
        print(f"[ERROR] Could not save Excel file: {e}")



class ScreenshotApp:
    def __init__(self, master):
        self.master = master
        master.title("Bulk Website Screenshot Tool")
        master.geometry("520x550")

        self.label = tk.Label(master, text="Manage your website list:")
        self.label.pack(pady=(10, 0))

        btn_frame = tk.Frame(master)
        btn_frame.pack(pady=5)

        self.write_urls_btn = tk.Button(btn_frame, text="Write URLs", command=self.open_url_window)
        self.write_urls_btn.grid(row=0, column=0, padx=5)

        self.load_csv_btn = tk.Button(btn_frame, text="Load CSV", command=self.load_csv)
        self.load_csv_btn.grid(row=0, column=1, padx=5)

        self.select_folder_btn = tk.Button(btn_frame, text="Select Output Folder", command=self.select_folder)
        self.select_folder_btn.grid(row=0, column=2, padx=5)

        label_format = tk.Label(master, text="Choose Output File Report Type:")
        label_format.pack(pady=(10,0))

        format_frame = tk.Frame(master)
        format_frame.pack(pady=5)
        self.output_format = tk.StringVar(value="Excel")
        rb_csv = tk.Radiobutton(format_frame, text="CSV", variable=self.output_format, value="csv")
        rb_excel = tk.Radiobutton(format_frame, text="Excel", variable=self.output_format, value="excel")
        rb_none = tk.Radiobutton(format_frame, text="None", variable=self.output_format, value="None")
        rb_csv.pack(side=tk.LEFT, padx=5)
        rb_excel.pack(side=tk.LEFT, padx=5)
        rb_none.pack(side=tk.LEFT, padx=5)

        action_frame = tk.Frame(master)
        action_frame.pack(pady=10)

        self.start_btn = tk.Button(action_frame, text="Start Screenshots", command=self.run_screenshot_thread)
        self.start_btn.grid(row=0, column=0, padx=5)

        self.stop_btn = tk.Button(action_frame, text="Stop", command=self.stop_screenshots, state='disabled')
        self.stop_btn.grid(row=0, column=1, padx=5)

        self.progress_label = tk.Label(master, text="Progress: 0%")
        self.progress_label.pack(pady=5)

        self.progress_bar = ttk.Progressbar(master, orient='horizontal', length=450, mode='determinate')
        self.progress_bar.pack(pady=5)

        self.progress_area = scrolledtext.ScrolledText(master, width=80, height=10, state='disabled')
        self.progress_area.pack(padx=10, pady=5)

        self.output_folder = os.path.expanduser("~/Desktop/Screenshot")
        self.headless_mode = True
        self.website_list = []
        self.stop_event = threading.Event()

    def open_url_window(self):
        popup = tk.Toplevel(self.master)
        popup.title("Enter URLs")
        popup.geometry("700x600")

        label = tk.Label(popup, text="Enter one URL per line:")
        label.pack(pady=5)

        self.popup_text = scrolledtext.ScrolledText(popup, width=50, height=15, bg='white', fg='black', font=('Helvetica', 12))
        self.popup_text.pack(padx=10, pady=5)

        def save_urls():
            raw_urls = self.popup_text.get("1.0", tk.END).splitlines()
            self.website_list = [url.strip() for url in raw_urls if url.strip()]
            self.log(f"Saved {len(self.website_list)} URLs.")
            popup.destroy()

        save_btn = tk.Button(popup, text="Save URLs", command=save_urls)
        save_btn.pack(pady=5)

    def load_csv(self):
        filepath = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if filepath:
            try:
                with open(filepath, 'r', newline='', encoding='utf-8') as csvfile:
                    reader = csv.reader(csvfile)
                    urls = []
                    for row in reader:
                        for item in row:
                            if item.strip():
                                urls.append(item.strip())
                    self.website_list = urls
                    self.log(f"Loaded {len(urls)} URLs from CSV.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load CSV file: {e}")

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_folder = folder
            self.log(f"Output folder set to: {folder}")

    def log(self, message):
        self.progress_area.config(state='normal')
        self.progress_area.insert(tk.END, message + "\n")
        self.progress_area.see(tk.END)
        self.progress_area.config(state='disabled')
        print(message)

    def run_screenshot_thread(self):
        self.stop_event.clear()
        self.start_btn.config(state='disabled')
        self.stop_btn.config(state='normal')
        thread = threading.Thread(target=self.start_screenshots)
        thread.start()

    def stop_screenshots(self):
        self.stop_event.set()
        self.log("Stop signal sent. Waiting for ongoing tasks to finish...")
        self.start_btn.config(state='normal')
        self.stop_btn.config(state='disabled')

    def start_screenshots(self):
        websites = self.website_list
        if not websites:
            messagebox.showwarning("Warning", "No websites provided.")
            self.start_btn.config(state='normal')
            self.stop_btn.config(state='disabled')
            return

        self.progress_bar['maximum'] = len(websites)
        self.progress_bar['value'] = 0
        self.log(f"Starting processing for {len(websites)} URLs...")
        total_urls = len(websites)
        successful_screenshots = 0
        url_results = []  

        max_workers = 10
        futures = []
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            for url in websites:
                if self.stop_event.is_set():
                    break
                futures.append(executor.submit(take_screenshot, url, self.output_folder, self.stop_event, self.headless_mode))

            for future in as_completed(futures):
                url_value, success, result_message, status_info = future.result()
                url_results.append((url_value, success, status_info))
                self.log(result_message)
                self.progress_bar['value'] += 1
                percentage = int((self.progress_bar['value'] / total_urls) * 100)
                self.progress_label.config(text=f"Progress: {percentage}%")
                if success:
                    successful_screenshots += 1

        if self.output_format.get() == "csv":
            save_results_to_csv(url_results, self.output_folder)
        elif self.output_format.get() == "excel":
            save_results_to_excel(url_results, self.output_folder)

        if self.stop_event.is_set():
            self.log("Screenshot process was stopped by user.")
        else:
            self.log("All screenshots processed.")

        self.show_complete_message(successful_screenshots, total_urls)
        self.start_btn.config(state='normal')
        self.stop_btn.config(state='disabled')

    def show_complete_message(self, successful_screenshots, total_urls):
        complete_popup = Toplevel(self.master)
        complete_popup.title("Process Complete")
        complete_popup.geometry("300x200")
        message = f"Process complete\n\nProcessed URLs: {total_urls}\nScreenshots Saved: {successful_screenshots}"
        label = tk.Label(complete_popup, text=message, font=("Helvetica", 10))
        label.pack(padx=10, pady=20)
        button = tk.Button(complete_popup, text="Close", command=complete_popup.destroy)
        button.pack(pady=10)



DRIVER_POOL = create_driver_pool(10, headless=True)


if __name__ == '__main__':
    root = tk.Tk()
    app = ScreenshotApp(root)
    root.mainloop()
