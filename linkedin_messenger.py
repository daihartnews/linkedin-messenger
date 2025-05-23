import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import time
import random
import json
import os
import csv
import threading
from datetime import datetime, timedelta
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, NoSuchElementException, NoSuchWindowException, WebDriverException
from fake_useragent import UserAgent
import psutil  # Optional for system monitoring

class LinkedInMessenger:
    def __init__(self, root):
        self.root = root
        self.root.title("LinkedIn Messenger")
        self.root.geometry("1000x700")
        self.driver = None
        self.contacts = []
        self.sort_column = None
        self.sort_reverse = False
        self.csv_file = "contacts.csv"
        self.json_file = "contacts.json"
        self.last_survey_file = "last_survey.txt"
        self.setup_gui()
        self.load_contacts()
        self.update_filter_suggestions()
        self.start_background_survey()

    def setup_gui(self):
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(fill="both", expand=True)

        left_pane = ttk.PanedWindow(main_pane, orient=tk.VERTICAL)
        right_pane = ttk.PanedWindow(main_pane, orient=tk.VERTICAL)
        main_pane.add(left_pane, weight=1)
        main_pane.add(right_pane, weight=1)

        q1_frame = ttk.LabelFrame(left_pane, text="Login & Filters", padding=10)
        left_pane.add(q1_frame, weight=1)

        ttk.Label(q1_frame, text="Email:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.email_entry = ttk.Entry(q1_frame, width=30)
        self.email_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(q1_frame, text="Password:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.password_entry = ttk.Entry(q1_frame, width=30, show="*")
        self.password_entry.grid(row=1, column=1, padx=5, pady=5)

        ttk.Button(q1_frame, text="Login", command=self.login_linkedin).grid(row=2, column=0, columnspan=2, pady=5)

        ttk.Label(q1_frame, text="Name:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.name_filter = ttk.Combobox(q1_frame, width=27, postcommand=self.update_filter_suggestions)
        self.name_filter.grid(row=3, column=1, padx=5, pady=5)

        ttk.Label(q1_frame, text="Job Title:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        self.job_filter = ttk.Combobox(q1_frame, width=27, postcommand=self.update_filter_suggestions)
        self.job_filter.grid(row=4, column=1, padx=5, pady=5)

        ttk.Label(q1_frame, text="Company:").grid(row=5, column=0, padx=5, pady=5, sticky="e")
        self.company_filter = ttk.Combobox(q1_frame, width=27, postcommand=self.update_filter_suggestions)
        self.company_filter.grid(row=5, column=1, padx=5, pady=5)

        ttk.Label(q1_frame, text="Industry:").grid(row=6, column=0, padx=5, pady=5, sticky="e")
        self.industry_filter = ttk.Combobox(q1_frame, width=27, postcommand=self.update_filter_suggestions)
        self.industry_filter.grid(row=6, column=1, padx=5, pady=5)

        button_frame = ttk.Frame(q1_frame)
        button_frame.grid(row=7, column=0, columnspan=2, pady=5)
        ttk.Button(button_frame, text="Search Contacts", command=self.search_contacts).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Save to CSV", command=self.save_contacts_to_csv).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Upload CSV", command=self.upload_csv).pack(side="left", padx=5)

        self.search_progress = ttk.Progressbar(q1_frame, mode="determinate")
        self.search_progress.grid(row=8, column=0, columnspan=2, pady=5, sticky="ew")

        q2_frame = ttk.LabelFrame(left_pane, text="Contacts", padding=10)
        left_pane.add(q2_frame, weight=2)

        tree_frame = ttk.Frame(q2_frame)
        tree_frame.pack(fill="both", expand=True)
        self.contacts_tree = ttk.Treeview(tree_frame, columns=("Select", "Name", "Job Title", "Company", "Industry"), show="headings")
        self.contacts_tree.heading("Select", text="Select")
        self.contacts_tree.heading("Name", text="Name", command=lambda: self.sort_contacts("Name"))
        self.contacts_tree.heading("Job Title", text="Job Title", command=lambda: self.sort_contacts("Job Title"))
        self.contacts_tree.heading("Company", text="Company", command=lambda: self.sort_contacts("Company"))
        self.contacts_tree.heading("Industry", text="Industry", command=lambda: self.sort_contacts("Industry"))
        self.contacts_tree.column("Select", width=50, anchor="center")
        self.contacts_tree.column("Name", width=150)
        self.contacts_tree.column("Job Title", width=150)
        self.contacts_tree.column("Company", width=150)
        self.contacts_tree.column("Industry", width=100)

        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.contacts_tree.yview)
        xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.contacts_tree.xview)
        self.contacts_tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.contacts_tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self.contacts_tree.bind("<Button-1>", self.handle_tree_click)

        button_frame = ttk.Frame(q2_frame)
        button_frame.pack(fill="x", pady=5)
        ttk.Button(button_frame, text="Select All", command=self.select_all_contacts).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Deselect All", command=self.deselect_all_contacts).pack(side="left", padx=5)

        q3_frame = ttk.LabelFrame(right_pane, text="Message", padding=10)
        right_pane.add(q3_frame, weight=1)

        ttk.Label(q3_frame, text="Selected Contacts:").grid(row=0, column=0, columnspan=2, padx=5, pady=5)
        self.selected_contacts_text = scrolledtext.ScrolledText(q3_frame, height=5, width=60, state="disabled")
        self.selected_contacts_text.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

        ttk.Label(q3_frame, text="Template:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.template_combo = ttk.Combobox(q3_frame, values=[
            "Hi {first_name}, I noticed your work at {company}. Let's connect!",
            "Hello {first_name}, I'm impressed by your role as {job_title}. Can we chat?",
            "Hi {first_name}, I'm in {industry} too. Let's discuss opportunities!"
        ], width=50)
        self.template_combo.grid(row=2, column=1, padx=5, pady=5)
        self.template_combo.bind("<<ComboboxSelected>>", self.load_template)

        ttk.Label(q3_frame, text="Message:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.message_text = scrolledtext.ScrolledText(q3_frame, height=5, width=60)
        self.message_text.grid(row=4, column=0, columnspan=2, padx=5, pady=5)

        q4_frame = ttk.LabelFrame(right_pane, text="Status", padding=10)
        right_pane.add(q4_frame, weight=2)

        ttk.Button(q4_frame, text="Send Messages", command=self.send_messages).pack(fill="x", pady=5)
        self.send_progress = ttk.Progressbar(q4_frame, mode="determinate")
        self.send_progress.pack(fill="x", pady=5)
        self.log_text = scrolledtext.ScrolledText(q4_frame, height=10, state="disabled")
        self.log_text.pack(fill="both", expand=True, pady=5)

    def log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")
        self.log_text.configure(state="disabled")
        self.log_text.see(tk.END)

    def log_system_stats(self):
        try:
            cpu_percent = psutil.cpu_percent(interval=1)
            memory = psutil.virtual_memory()
            memory_percent = memory.percent
            chrome_process = None
            for proc in psutil.process_iter(['name']):
                if proc.info['name'].lower() in ['chrome.exe', 'chromedriver.exe']:
                    chrome_process = proc
                    break
            chrome_memory = chrome_process.memory_info().rss / 1024 / 1024 if chrome_process else 0
            self.log(f"System stats: CPU={cpu_percent}%, Memory={memory_percent}%, Chrome Memory={chrome_memory:.2f} MB")
        except Exception as e:
            self.log(f"Failed to log system stats: {str(e)}")

    def load_contacts(self):
        if os.path.exists(self.json_file):
            try:
                with open(self.json_file, "r", encoding="utf-8") as file:
                    data = json.load(file)
                    self.contacts = [{"name": d["name"], "job_title": d.get("job_title", ""), "company": d.get("company", ""), "industry": d.get("industry", ""), "element": None, "selected": False} for d in data]
                self.log(f"Loaded {len(self.contacts)} contacts from {self.json_file}")
            except Exception as e:
                self.log(f"Error loading from JSON: {str(e)}")
        elif os.path.exists(self.csv_file):
            try:
                with open(self.csv_file, mode="r", encoding="utf-8") as file:
                    reader = csv.DictReader(file)
                    self.contacts = []
                    for row in reader:
                        contact = {
                            "name": row["name"],
                            "job_title": row["job_title"],
                            "company": row["company"],
                            "industry": row["industry"],
                            "element": None,
                            "selected": False
                        }
                        self.contacts.append(contact)
                self.log(f"Loaded {len(self.contacts)} contacts from {self.csv_file}")
            except Exception as e:
                self.log(f"Error loading from CSV: {str(e)}")
        else:
            self.log("No contacts file found")

    def save_contacts_to_json(self):
        try:
            with open(self.json_file, "w", encoding="utf-8") as file:
                json.dump([{"name": c["name"], "job_title": c["job_title"], "company": c["company"], "industry": c["industry"]} for c in self.contacts], file, ensure_ascii=False, indent=4)
            self.log(f"Saved {len(self.contacts)} contacts to {self.json_file}")
        except Exception as e:
            self.log(f"Error saving to JSON: {str(e)}")

    def save_contacts_to_csv(self):
        try:
            with open(self.csv_file, mode="w", encoding="utf-8", newline="") as file:
                writer = csv.DictWriter(file, fieldnames=["name", "job_title", "company", "industry"])
                writer.writeheader()
                for contact in self.contacts:
                    writer.writerow({
                        "name": contact["name"],
                        "job_title": contact["job_title"],
                        "company": contact["company"],
                        "industry": contact["industry"]
                    })
            self.log(f"Saved {len(self.contacts)} contacts to {self.csv_file}")
        except Exception as e:
            self.log(f"Error saving to CSV: {str(e)}")

    def upload_csv(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if file_path:
            try:
                self.contacts = []
                with open(file_path, mode="r", encoding="utf-8") as file:
                    reader = csv.DictReader(file)
                    for row in reader:
                        contact = {
                            "name": row["name"],
                            "job_title": row["job_title"],
                            "company": row["company"],
                            "industry": row["industry"],
                            "element": None,
                            "selected": False
                        }
                        self.contacts.append(contact)
                self.log(f"Uploaded {len(self.contacts)} contacts from {file_path}")
                self.display_contacts(self.contacts)
                self.update_filter_suggestions()
            except Exception as e:
                self.log(f"Error uploading CSV: {str(e)}")
                messagebox.showerror("Error", f"Failed to upload CSV: {str(e)}")

    def update_filter_suggestions(self):
        names = sorted(set(c["name"] for c in self.contacts))
        job_titles = sorted(set(c["job_title"] for c in self.contacts if c["job_title"]))
        companies = sorted(set(c["company"] for c in self.contacts if c["company"]))
        industries = sorted(set(c["industry"] for c in self.contacts if c["industry"]))
        self.name_filter["values"] = names
        self.job_filter["values"] = job_titles
        self.company_filter["values"] = companies
        self.industry_filter["values"] = industries

    def login_linkedin(self):
        email = self.email_entry.get()
        password = self.password_entry.get()
        if not email or not password:
            messagebox.showerror("Error", "Please enter email and password")
            return

        def login():
            try:
                ua = UserAgent()
                options = webdriver.ChromeOptions()
                options.add_argument(f"user-agent={ua.random}")
                options.add_argument("--disable-blink-features=AutomationControlled")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-dev-shm-usage")
                options.add_argument("--disable-gpu")
                options.add_argument("--disable-web-security")
                options.add_argument("--disable-notifications")
                options.add_argument("--disable-popup-blocking")
                options.add_argument("--window-size=1280,720")
                options.add_experimental_option("excludeSwitches", ["enable-automation"])
                options.add_experimental_option("useAutomationExtension", False)
                self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
                self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                    "source": """
                        Object.defineProperty(navigator, 'webdriver', {
                            get: () => undefined
                        });
                        Object.defineProperty(navigator, 'plugins', {
                            get: () => [1, 2, 3]
                        });
                        Object.defineProperty(navigator, 'languages', {
                            get: () => ['en-US', 'en']
                        });
                        window.chrome = window.chrome || {};
                        window.chrome.runtime = {};
                    """
                })
                self.log("Navigating to LinkedIn login page")
                self.driver.set_page_load_timeout(30)
                self.driver.get("https://www.linkedin.com/login")
                time.sleep(random.uniform(2, 4))

                email_field = self.driver.find_element(By.ID, "username")
                email_field.send_keys(email)
                password_field = self.driver.find_element(By.ID, "password")
                password_field.send_keys(password)
                password_field.send_keys(Keys.RETURN)
                time.sleep(random.uniform(5, 7))

                if "login" in self.driver.current_url:
                    self.root.after(0, lambda: self.log("Login failed: Check credentials"))
                    self.root.after(0, lambda: messagebox.showerror("Error", "Login failed. Check credentials."))
                    self.driver.quit()
                    self.driver = None
                    return

                self.root.after(0, lambda: self.log("Logged in successfully"))
            except Exception as e:
                self.root.after(0, lambda: self.log(f"Login error: {str(e)}"))
                self.root.after(0, lambda: messagebox.showerror("Error", f"Login failed: {str(e)}"))
                if self.driver:
                    self.driver.quit()
                    self.driver = None

        threading.Thread(target=login, daemon=True).start()

    def search_contacts(self):
        filtered_contacts = self.apply_filters(self.contacts)
        self.search_progress["maximum"] = len(self.contacts) or 1
        self.search_progress["value"] = 0

        if filtered_contacts:
            self.log(f"Found {len(filtered_contacts)} matching contacts locally")
            self.display_contacts(filtered_contacts)
            self.search_progress["value"] = len(self.contacts)
            return

        if not self.driver:
            messagebox.showerror("Error", "Please log in first")
            return

        def search():
            try:
                self.root.after(0, lambda: self.contacts_tree.delete(*self.contacts_tree.get_children()))
                new_contacts = []
                self.driver.set_page_load_timeout(30)
                self.driver.get("https://www.linkedin.com/mynetwork/invite-connect/connections/")
                time.sleep(random.uniform(5, 7))

                last_height = self.driver.execute_script("return document.body.scrollHeight")
                while True:
                    self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(random.uniform(3, 5))
                    new_height = self.driver.execute_script("return document.body.scrollHeight")
                    if new_height == last_height:
                        break
                    last_height = new_height

                contact_elements = self.driver.find_elements(By.CSS_SELECTOR, ".mn-connection-card, .connection-card")
                existing_names = {c["name"] for c in self.contacts}

                for elem in contact_elements:
                    try:
                        name_elem = elem.find_element(By.CSS_SELECTOR, "[class*='name'], .t-16.t-black.t-bold")
                        name = name_elem.text.strip() if name_elem else ""
                        details_elem = elem.find_element(By.CSS_SELECTOR, "[class*='occupation'], .t-14.t-black--light")
                        details = details_elem.text.strip() if details_elem else ""
                        job_title = details.split(" at ")[0].strip() if " at " in details else details
                        company = details.split(" at ")[1].strip() if " at " in details else ""
                        industry = ""
                        if name and name not in existing_names:
                            contact = {
                                "name": name,
                                "job_title": job_title,
                                "company": company,
                                "industry": industry,
                                "element": elem,
                                "selected": False
                            }
                            new_contacts.append(contact)
                            existing_names.add(name)
                    except Exception as e:
                        self.root.after(0, lambda: self.log(f"Skipped contact due to error: {str(e)}"))
                        continue

                self.contacts.extend(new_contacts)
                self.root.after(0, lambda: self.save_contacts_to_json())
                filtered_contacts = self.apply_filters(self.contacts)
                self.root.after(0, lambda: self.display_contacts(filtered_contacts))
                self.root.after(0, lambda: self.log(f"Searched {len(self.contacts)} total contacts, displayed {len(filtered_contacts)} after filtering"))
                self.root.after(0, lambda: self.update_filter_suggestions())
            except Exception as e:
                self.root.after(0, lambda: self.log(f"Error searching contacts: {str(e)}"))
            finally:
                self.root.after(0, lambda: self.search_progress.configure(value=len(self.contacts)))

        threading.Thread(target=search, daemon=True).start()

    def apply_filters(self, contacts):
        name_filters = [n.strip() for n in self.name_filter.get().lower().split(";") if n.strip()]
        job_filters = [j.strip() for j in self.job_filter.get().lower().split(";") if j.strip()]
        company_filters = [c.strip() for c in self.company_filter.get().lower().split(";") if c.strip()]
        industry_filters = [i.strip() for i in self.industry_filter.get().lower().split(";") if i.strip()]

        filtered_contacts = []
        self.search_progress["maximum"] = len(contacts) or 1
        for i, contact in enumerate(contacts):
            self.search_progress["value"] = i + 1
            self.root.update()
            if name_filters and not any(n in contact["name"].lower() for n in name_filters):
                continue
            if job_filters and not any(j in contact["job_title"].lower() for j in job_filters):
                continue
            if company_filters and not any(c in contact["company"].lower() for c in company_filters):
                continue
            if industry_filters and not any(i in contact["industry"].lower() for i in industry_filters):
                continue
            filtered_contacts.append(contact)
        self.search_progress["value"] = len(contacts)
        return filtered_contacts

    def display_contacts(self, contacts):
        self.contacts_tree.delete(*self.contacts_tree.get_children())
        for contact in contacts:
            self.contacts_tree.insert("", "end", iid=contact["name"], values=(
                "☑" if contact["selected"] else "☐",
                contact["name"],
                contact["job_title"],
                contact["company"],
                contact["industry"]
            ))
        self.update_selected_contacts()

    def start_background_survey(self):
        def survey():
            while True:
                if self.driver and self.should_survey():
                    self.log("Starting background LinkedIn contact survey")
                    self.survey_linkedin_contacts()
                time.sleep(24 * 3600)  # Check daily

        threading.Thread(target=survey, daemon=True).start()

    def should_survey(self):
        if not os.path.exists(self.last_survey_file):
            return True
        try:
            with open(self.last_survey_file, "r") as file:
                last_survey = datetime.fromisoformat(file.read().strip())
            return datetime.now() - last_survey >= timedelta(days=14)
        except Exception as e:
            self.log(f"Error reading survey timestamp: {str(e)}")
            return True

    def survey_linkedin_contacts(self):
        try:
            self.driver.set_page_load_timeout(30)
            self.driver.get("https://www.linkedin.com/mynetwork/invite-connect/connections/")
            time.sleep(random.uniform(5, 7))

            last_height = self.driver.execute_script("return document.body.scrollHeight")
            while True:
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(random.uniform(3, 5))
                new_height = self.driver.execute_script("return document.body.scrollHeight")
                if new_height == last_height:
                    break
                last_height = new_height

            contact_elements = self.driver.find_elements(By.CSS_SELECTOR, ".mn-connection-card, .connection-card")
            existing_names = {c["name"] for c in self.contacts}
            new_contacts = []

            for elem in contact_elements:
                try:
                    name_elem = elem.find_element(By.CSS_SELECTOR, "[class*='name'], .t-16.t-black.t-bold")
                    name = name_elem.text.strip() if name_elem else ""
                    details_elem = elem.find_element(By.CSS_SELECTOR, "[class*='occupation'], .t-14.t-black--light")
                    details = details_elem.text.strip() if details_elem else ""
                    job_title = details.split(" at ")[0].strip() if " at " in details else details
                    company = details.split(" at ")[1].strip() if " at " in details else ""
                    industry = ""
                    if name and name not in existing_names:
                        contact = {
                            "name": name,
                            "job_title": job_title,
                            "company": company,
                            "industry": industry,
                            "element": None,
                            "selected": False
                        }
                        new_contacts.append(contact)
                        existing_names.add(name)
                except Exception as e:
                    self.log(f"Skipped contact in survey due to error: {str(e)}")
                    continue

            if new_contacts:
                self.contacts.extend(new_contacts)
                self.save_contacts_to_json()
                self.log(f"Survey added {len(new_contacts)} new contacts")
            else:
                self.log("Survey found no new contacts")

            try:
                with open(self.last_survey_file, "w") as file:
                    file.write(datetime.now().isoformat())
            except Exception as e:
                self.log(f"Error writing survey timestamp: {str(e)}")
        except Exception as e:
            self.log(f"Survey error: {str(e)}")

    def handle_tree_click(self, event):
        item = self.contacts_tree.identify_row(event.y)
        column = self.contacts_tree.identify_column(event.x)
        if item and column == "#1":
            contact_name = self.contacts_tree.item(item)["values"][1]
            contact = next(c for c in self.contacts if c["name"] == contact_name)
            contact["selected"] = not contact["selected"]
            self.contacts_tree.item(item, values=(
                "☑" if contact["selected"] else "☐",
                contact["name"],
                contact["job_title"],
                contact["company"],
                contact["industry"]
            ))
            self.update_selected_contacts()

    def sort_contacts(self, column):
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_reverse = False
            self.sort_column = column

        col_map = {
            "Name": "name",
            "Job Title": "job_title",
            "Company": "company",
            "Industry": "industry"
        }
        key = col_map[column]

        sorted_contacts = sorted(self.contacts, key=lambda x: x[key].lower() if x[key] else "", reverse=self.sort_reverse)
        self.display_contacts(sorted_contacts)

    def select_all_contacts(self):
        for contact in self.contacts:
            contact["selected"] = True
        self.display_contacts(self.contacts)

    def deselect_all_contacts(self):
        for contact in self.contacts:
            contact["selected"] = False
        self.display_contacts(self.contacts)

    def load_template(self, event):
        self.message_text.delete("1.0", tk.END)
        self.message_text.insert("1.0", self.template_combo.get())
        self.update_selected_contacts()

    def update_selected_contacts(self):
        self.selected_contacts_text.configure(state="normal")
        self.selected_contacts_text.delete("1.0", tk.END)
        selected_contacts = [c for c in self.contacts if c["selected"]]

        if selected_contacts:
            self.selected_contacts_text.insert(tk.END, "Selected Contacts:\n")
            for contact in selected_contacts:
                self.selected_contacts_text.insert(tk.END, f"- {contact['name']}\n")
        else:
            self.selected_contacts_text.insert(tk.END, "No contacts selected.\n")

        self.selected_contacts_text.configure(state="disabled")
        self.selected_contacts_text.see(tk.END)

    def check_page_state(self):
        try:
            self.log("Checking page state")
            # Wait for basic HTML to load
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "body"))
            )
            self.log("Page body loaded")

            # Check JavaScript readiness
            try:
                WebDriverWait(self.driver, 10).until(
                    lambda d: d.execute_script("return document.readyState") == "complete"
                )
                self.log("Page JavaScript completed")
            except TimeoutException:
                self.log("Warning: Page JavaScript did not complete, forcing load")
                self.driver.execute_script("window.stop();")  # Force stop loading
                # Save HTML for debugging
                try:
                    html_snippet = self.driver.page_source[:1000]
                    html_path = f"html_page_load_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
                    with open(html_path, "w", encoding="utf-8") as f:
                        f.write(html_snippet)
                    self.log(f"Saved page HTML: {html_path}")
                except Exception as e:
                    self.log(f"Failed to save page HTML: {str(e)}")

            page_url = self.driver.current_url
            page_title = self.driver.title
            captcha_detected = bool(self.driver.find_elements(By.CSS_SELECTOR, "[id*='captcha'], [class*='captcha'], iframe[src*='captcha']"))
            restricted = bool(self.driver.find_elements(By.XPATH, "//h1[contains(text(), 'restricted') or contains(text(), 'suspended')]"))
            rate_limit = bool(self.driver.find_elements(By.XPATH, "//*[contains(text(), 'weekly limit') or contains(text(), 'reached the limit') or contains(text(), 'too many requests')]"))
            browser_logs = self.driver.get_log("browser") if hasattr(self.driver, "get_log") else []

            self.log(f"Page state: URL={page_url}, Title={page_title}, CAPTCHA={captcha_detected}, Restricted={restricted}, RateLimit={rate_limit}")
            if browser_logs:
                self.log(f"Browser console errors: {browser_logs}")

            if "login" in page_url.lower():
                self.log("Session expired, login required")
                self.root.after(0, lambda: messagebox.showerror("Error", "Session expired. Please log in again."))
                return False
            if captcha_detected:
                self.log("CAPTCHA detected. Please resolve manually.")
                self.root.after(0, lambda: messagebox.showinfo("CAPTCHA", "Please resolve the CAPTCHA in the browser and press OK to continue."))
                WebDriverWait(self.driver, 60).until_not(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "[id*='captcha'], [class*='captcha'], iframe[src*='captcha']"))
                )
                self.log("CAPTCHA resolved, continuing")
                return True
            if restricted:
                self.log("Account restricted by LinkedIn")
                self.root.after(0, lambda: messagebox.showerror("Error", "Account restricted. Please resolve and try again."))
                return False
            if rate_limit:
                self.log("Rate limit detected. Please wait and try again later.")
                self.root.after(0, lambda: messagebox.showwarning("Rate Limit", "LinkedIn has imposed a rate limit. Please wait and try again later."))
                return False
            return True
        except Exception as e:
            self.log(f"Error checking page state: {str(e)}")
            # Save screenshot and HTML
            try:
                screenshot_path = f"screenshot_page_state_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                self.driver.save_screenshot(screenshot_path)
                self.log(f"Saved screenshot: {screenshot_path}")
            except Exception as e:
                self.log(f"Failed to save screenshot: {str(e)}")
            try:
                html_snippet = self.driver.page_source[:1000]
                html_path = f"html_page_state_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
                with open(html_path, "w", encoding="utf-8") as f:
                        f.write(html_snippet)
                self.log(f"Saved page HTML: {html_path}")
            except Exception as e:
                self.log(f"Failed to save page HTML: {str(e)}")
            return False

    def send_messages(self):
        selected_contacts = [c for c in self.contacts if c["selected"]]
        if not selected_contacts:
            self.log("No contacts selected for messaging")
            return

        message = self.message_text.get("1.0", tk.END).strip()
        if not message:
            self.log("No message entered for sending")
            return

        if not self.driver:
            messagebox.showerror("Error", "Please log in first")
            return

        self.send_progress["maximum"] = len(selected_contacts)
        self.send_progress["value"] = 0
        successful_sends = 0

        def send():
            nonlocal successful_sends
            try:
                self.log("Starting message sending process")
                self.log_system_stats()
                # Save initial screenshot
                try:
                    screenshot_path = f"screenshot_pre_nav_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                    self.driver.save_screenshot(screenshot_path)
                    self.log(f"Saved pre-navigation screenshot: {screenshot_path}")
                except Exception as e:
                    self.log(f"Failed to save pre-navigation screenshot: {str(e)}")

                # Pre-navigation refresh
                self.log("Refreshing browser before navigation")
                self.driver.refresh()
                time.sleep(random.uniform(2, 4))

                # Try connections page with retries
                self.log("Attempting to navigate to connections page")
                self.driver.set_page_load_timeout(30)
                primary_url = "https://www.linkedin.com/mynetwork/invite-connect/connections/"
                fallback_url = "https://www.linkedin.com/messaging/"
                navigation_success = False

                for attempt in range(3):  # Increased to 3 retries
                    try:
                        self.log(f"Navigation attempt {attempt + 1}/3 to {primary_url}")
                        self.driver.get(primary_url)
                        time.sleep(random.uniform(5, 7))
                        navigation_success = True
                        break
                    except TimeoutException as e:
                        self.log(f"Navigation timeout (attempt {attempt + 1}/3): {str(e)}")
                        if attempt < 2:
                            self.log("Retrying navigation after refresh")
                            self.driver.refresh()
                            time.sleep(random.uniform(5, 7))
                        continue

                # Try fallback URL if primary fails
                if not navigation_success:
                    self.log(f"Primary navigation failed, attempting fallback: {fallback_url}")
                    for attempt in range(2):
                        try:
                            self.log(f"Fallback navigation attempt {attempt + 1}/2 to {fallback_url}")
                            self.driver.get(fallback_url)
                            time.sleep(random.uniform(5, 7))
                            navigation_success = True
                            break
                        except TimeoutException as e:
                            self.log(f"Fallback navigation timeout (attempt {attempt + 1}/2): {str(e)}")
                            if attempt == 0:
                                self.log("Retrying fallback navigation after refresh")
                                self.driver.refresh()
                                time.sleep(random.uniform(5, 7))
                            continue

                if not navigation_success:
                    self.log("All navigation attempts failed. Aborting.")
                    self.root.after(0, lambda: messagebox.showerror("Error", "Failed to load LinkedIn page. Check network or LinkedIn status."))
                    return

                self.log("Checking page state after navigation")
                if not self.check_page_state():
                    self.log("Failed to load valid page state. Aborting.")
                    self.root.after(0, lambda: messagebox.showerror("Error", "Invalid page state. Check for CAPTCHAs or restrictions."))
                    return

                for i, contact in enumerate(selected_contacts):
                    max_retries = 2
                    for retry in range(max_retries):
                        try:
                            self.log(f"Processing {contact['name']} (retry {retry + 1}/{max_retries})")
                            if not self.check_page_state():
                                self.log(f"Skipping {contact['name']}: Invalid page state")
                                break

                            # Check if browser window is still open
                            try:
                                self.driver.current_url
                            except NoSuchWindowException:
                                self.log("Browser window closed. Reinitializing WebDriver.")
                                self.driver.quit()
                                self.driver = None
                                ua = UserAgent()
                                options = webdriver.ChromeOptions()
                                options.add_argument(f"user-agent={ua.random}")
                                options.add_argument("--disable-blink-features=AutomationControlled")
                                options.add_argument("--no-sandbox")
                                options.add_argument("--disable-dev-shm-usage")
                                options.add_argument("--disable-gpu")
                                options.add_argument("--disable-web-security")
                                options.add_argument("--disable-notifications")
                                options.add_argument("--disable-popup-blocking")
                                options.add_argument("--window-size=1280,720")
                                options.add_experimental_option("excludeSwitches", ["enable-automation"])
                                options.add_experimental_option("useAutomationExtension", False)
                                self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
                                self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                                    "source": """
                                        Object.defineProperty(navigator, 'webdriver', {
                                            get: () => undefined
                                        });
                                        Object.defineProperty(navigator, 'plugins', {
                                            get: () => [1, 2, 3]
                                        });
                                        Object.defineProperty(navigator, 'languages', {
                                            get: () => ['en-US', 'en']
                                        });
                                        window.chrome = window.chrome || {};
                                        window.chrome.runtime = {};
                                    """
                                })
                                self.driver.set_page_load_timeout(30)
                                self.driver.get(primary_url)
                                time.sleep(random.uniform(5, 7))
                                if not self.check_page_state():
                                    self.log(f"Skipping {contact['name']}: Failed to reload connections page after reinitialization")
                                    break

                            # Pre-wait for page stability
                            time.sleep(random.uniform(2, 4))

                            # Search for contact with retries
                            contact_element = None
                            for attempt in range(3):
                                self.log(f"Searching for {contact['name']} (attempt {attempt + 1}/3)")
                                try:
                                    # Try multiple selectors for search input
                                    selectors = [
                                        (By.CSS_SELECTOR, "input[placeholder*='Search'][type='text'], [id*='search'], [class*='search-input'], [id*='connections-search']"),
                                        (By.XPATH, "//input[contains(@placeholder, 'Search') or contains(@id, 'search') or contains(@class, 'search')]"),
                                        (By.CSS_SELECTOR, "input.search-global-typeahead__input, input.global-nav__search-input"),
                                        (By.XPATH, "//input[@role='combobox' and contains(@class, 'search')]")
                                    ]
                                    search_input = None
                                    for by_type, selector in selectors:
                                        try:
                                            search_input = WebDriverWait(self.driver, 10).until(
                                                EC.presence_of_element_located((by_type, selector))
                                            )
                                            self.log(f"Search input found for {contact['name']} with {by_type}: {selector}")
                                            break
                                        except TimeoutException:
                                            self.log(f"Selector failed for {contact['name']}: {by_type} - {selector}")
                                            continue

                                    if not search_input:
                                        raise TimeoutException("No search input found with any selector")

                                    self.log(f"Search input HTML for {contact['name']}: {search_input.get_attribute('outerHTML')[:100]}")
                                    search_input.clear()
                                    time.sleep(random.uniform(0.5, 1))
                                    search_input.clear()  # Double clear to ensure empty
                                    for char in contact["name"]:
                                        search_input.send_keys(char)
                                        time.sleep(random.uniform(0.05, 0.15))
                                    self.log(f"Entered search query: {contact['name']}")
                                    search_input.send_keys(Keys.RETURN)
                                    time.sleep(random.uniform(5, 7))

                                    # Log current URL and title
                                    self.log(f"Post-search state: URL={self.driver.current_url}, Title={self.driver.title}")

                                    # Handle potential redirection
                                    if "search/results/all" in self.driver.current_url:
                                        self.log(f"Redirected to global search, navigating back for {contact['name']}")
                                        self.driver.get(primary_url)
                                        time.sleep(random.uniform(5, 7))
                                        if not self.check_page_state():
                                            self.log(f"Skipping {contact['name']}: Failed to reload connections page")
                                            continue
                                        continue

                                    # Wait for search results
                                    try:
                                        results_container = WebDriverWait(self.driver, 10).until(
                                            EC.presence_of_element_located((By.CSS_SELECTOR, "ul[class*='connections'], div[class*='search-results'], div[class*='entity-result']"))
                                        )
                                        self.log(f"Search results container found for {contact['name']}")
                                        contact_elements = self.driver.find_elements(By.CSS_SELECTOR, "li[class*='search-result'], li.reusable-search__result-container, .mn-connection-card, .connection-card, div[class*='entity-result'], div[class*='search-result']")
                                        self.log(f"Found {len(contact_elements)} contact elements for {contact['name']}")
                                        for elem in contact_elements:
                                            try:
                                                name_elem = elem.find_element(By.CSS_SELECTOR, "[class*='title'], [class*='name'], h3, h4, .t-16, .t-14, span")
                                                name_text = name_elem.text.strip()
                                                self.log(f"Checking contact: {name_text}")
                                                if contact["name"].lower() in name_text.lower():
                                                    contact_element = elem
                                                    self.log(f"Found {contact['name']} in search results")
                                                    break
                                            except (StaleElementReferenceException, NoSuchElementException):
                                                continue
                                        if contact_element:
                                            break
                                    except Exception as e:
                                        self.log(f"Search attempt {attempt + 1} failed for {contact['name']}: {str(e)}")
                                        continue
                                except TimeoutException as e:
                                    self.log(f"Timeout finding search input for {contact['name']} (attempt {attempt + 1}): {str(e)}")
                                    # Save HTML snippet for debugging
                                    try:
                                        html_snippet = self.driver.page_source[:1000]
                                        with open(f"html_{contact['name'].replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html", "w", encoding="utf-8") as f:
                                            f.write(html_snippet)
                                        self.log(f"Saved HTML snippet: html_{contact['name'].replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
                                    except Exception as e:
                                        self.log(f"Failed to save HTML snippet for {contact['name']}: {str(e)}")
                                    # Refresh page on last attempt
                                    if attempt == 2:
                                        self.log(f"Refreshing page for {contact['name']} after failed attempts")
                                        self.driver.refresh()
                                        time.sleep(random.uniform(5, 7))
                                        if not self.check_page_state():
                                            self.log(f"Skipping {contact['name']}: Failed to reload page after refresh")
                                            break
                                    continue
                                except NoSuchElementException as e:
                                    self.log(f"Search input not found for {contact['name']} (attempt {attempt + 1}): {str(e)}")
                                    continue

                            if not contact_element:
                                self.log(f"Skipping {contact['name']}: Contact not found in search results after 3 attempts")
                                # Take screenshot for debugging
                                try:
                                    screenshot_path = f"screenshot_{contact['name'].replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                                    self.driver.save_screenshot(screenshot_path)
                                    self.log(f"Saved screenshot: {screenshot_path}")
                                except Exception as e:
                                    self.log(f"Failed to save screenshot for {contact['name']}: {str(e)}")
                                continue

                            # Log contact element HTML
                            try:
                                contact_html = contact_element.get_attribute("outerHTML")[:500]
                                self.log(f"Contact element HTML for {contact['name']}: {contact_html}")
                            except Exception as e:
                                self.log(f"Failed to get contact element HTML for {contact['name']}: {str(e)}")

                            # Scroll and hover to contact element
                            self.driver.execute_script("arguments[0].scrollIntoView(true);", contact_element)
                            time.sleep(random.uniform(0.3, 0.6))
                            ActionChains(self.driver).move_to_element(contact_element).perform()
                            time.sleep(random.uniform(0.3, 0.6))

                            # Debug: Log all buttons in contact element
                            buttons = contact_element.find_elements(By.CSS_SELECTOR, "button")
                            button_texts = [(btn.text.strip(), btn.get_attribute("aria-label") or "", btn.get_attribute("data-control-name") or "") for btn in buttons]
                            self.log(f"Buttons in contact element for {contact['name']}: {button_texts}")

                            # Try clicking 'More' button if present
                            try:
                                more_button = contact_element.find_element(By.CSS_SELECTOR, "button[aria-label*='More actions' i], button[class*='more']")
                                self.driver.execute_script("arguments[0].scrollIntoView(true);", more_button)
                                ActionChains(self.driver).move_to_element(more_button).click().perform()
                                self.log(f"Clicked 'More' button for {contact['name']}")
                                time.sleep(random.uniform(1, 2))
                            except (NoSuchElementException, TimeoutException):
                                self.log(f"No 'More' button found for {contact['name']}")

                            # Try multiple selectors for message button
                            message_button = None
                            successful_selectors = []
                            failed_selectors = []

                            # Define selector strategies
                            selectors = [
                                # CSS Selectors (contact element)
                                {"selector": "button[class*='message'], button[class*='msg'], button[class*='artdeco-button'], button[class*='pv-top-card']", "context": "contact element", "type": "css"},
                                {"selector": "button[aria-label*='message' i], button[aria-label*='Send a message' i]", "context": "contact element", "type": "css"},
                                {"selector": "[data-control-name*='message']", "context": "contact element", "type": "css"},
                                # XPath Selectors (contact element)
                                {"selector": ".//button[contains(text(), 'Message') or contains(text(), 'Send a message')]", "context": "contact element", "type": "xpath"},
                                {"selector": ".//button[contains(@aria-label, 'Message') or contains(@aria-label, 'Send a message')]", "context": "contact element", "type": "xpath"},
                                {"selector": ".//button[@data-control-name='message']", "context": "contact element", "type": "xpath"},
                                # CSS Selectors (page-wide)
                                {"selector": f"button[aria-label*='{contact['name']}' i][aria-label*='message' i], button[aria-label*='{contact['name']}' i][aria-label*='Send a message' i]", "context": "page-wide", "type": "css"},
                                {"selector": "button[class*='message'], button[class*='msg'], button[class*='artdeco-button'], button[class*='pv-top-card']", "context": "page-wide", "type": "css"},
                                {"selector": "button[aria-label*='message' i], button[aria-label*='Send a message' i]", "context": "page-wide", "type": "css"},
                                {"selector": "[data-control-name*='message']", "context": "page-wide", "type": "css"},
                                # XPath Selectors (page-wide)
                                {"selector": "//button[contains(text(), 'Message') or contains(text(), 'Send a message')]", "context": "page-wide", "type": "xpath"},
                                {"selector": "//button[contains(@aria-label, 'Message') or contains(@aria-label, 'Send a message')]", "context": "page-wide", "type": "xpath"},
                                {"selector": "//button[@data-control-name='message']", "context": "page-wide", "type": "xpath"}
                            ]

                            # Try each selector with retries
                            for sel in selectors:
                                for attempt in range(2):
                                    try:
                                        parent = contact_element if sel["context"] == "contact element" else self.driver
                                        by_type = By.CSS_SELECTOR if sel["type"] == "css" else By.XPATH
                                        button = WebDriverWait(parent, 10).until(
                                            EC.visibility_of_element_located((by_type, sel["selector"]))
                                        )
                                        WebDriverWait(parent, 10).until(
                                            EC.element_to_be_clickable((by_type, sel["selector"]))
                                        )
                                        self.log(f"Found message button for {contact['name']} with selector '{sel['selector']}' in {sel['context']} (attempt {attempt + 1})")
                                        successful_selectors.append(f"{sel['selector']} ({sel['context']})")
                                        if not message_button:  # Use the first successful button
                                            message_button = button
                                        break
                                    except (TimeoutException, NoSuchElementException) as e:
                                        self.log(f"Failed to find message button for {contact['name']} with selector '{sel['selector']}' in {sel['context']} (attempt {attempt + 1}): {str(e)}")
                                        failed_selectors.append(f"{sel['selector']} ({sel['context']})")
                                        time.sleep(random.uniform(1, 2))
                                        continue

                            # Try profile navigation approach
                            try:
                                self.log(f"Trying profile navigation for {contact['name']}")
                                name_link = contact_element.find_element(By.CSS_SELECTOR, "a[href*='/in/'], a[class*='app-aware-link'], a[class*='entity']")
                                self.driver.execute_script("arguments[0].scrollIntoView(true);", name_link)
                                time.sleep(random.uniform(0.3, 0.6))
                                ActionChains(self.driver).move_to_element(name_link).click().perform()
                                time.sleep(random.uniform(3, 5))
                                if not self.check_page_state():
                                    self.log(f"Skipping profile navigation for {contact['name']}: Invalid page state")
                                    raise TimeoutException("Invalid page state")
                                message_button_profile = WebDriverWait(self.driver, 10).until(
                                    EC.element_to_be_clickable((By.CSS_SELECTOR, "button[class*='message'], button[class*='msg'], button[class*='artdeco-button'], button[class*='pv-top-card'], button[aria-label*='message' i], button[aria-label*='Send a message' i]"))
                                )
                                self.log(f"Found message button for {contact['name']} via profile navigation")
                                successful_selectors.append("profile navigation")
                                if not message_button:  # Use profile button if no prior success
                                    message_button = message_button_profile
                                # Navigate back to connections page
                                self.driver.get(primary_url)
                                time.sleep(random.uniform(5, 7))
                            except (TimeoutException, NoSuchElementException, StaleElementReferenceException) as e:
                                self.log(f"Failed profile navigation for {contact['name']}: {str(e)}")
                                failed_selectors.append("profile navigation")
                                # Navigate back to connections page
                                self.driver.get(primary_url)
                                time.sleep(random.uniform(5, 7))

                            # Summarize selector results
                            self.log(f"Selector summary for {contact['name']}: Successful: {successful_selectors}; Failed: {failed_selectors}")

                            if not message_button:
                                self.log(f"Skipping {contact['name']}: No message button found with any selector")
                                continue

                            # Click message button
                            self.driver.execute_script("arguments[0].scrollIntoView(true);", message_button)
                            time.sleep(random.uniform(0.3, 0.6))
                            ActionChains(self.driver).move_to_element(message_button).click().perform()
                            self.log(f"Clicked message button for {contact['name']}")
                            time.sleep(random.uniform(1, 2))

                            # Send message
                            first_name = contact["name"].split()[0]
                            formatted_message = message.format(
                                first_name=first_name,
                                name=contact["name"],
                                job_title=contact["job_title"],
                                company=contact["company"],
                                industry=contact["industry"]
                            )
                            message_input = WebDriverWait(self.driver, 10).until(
                                EC.presence_of_element_located((By.CSS_SELECTOR, ".msg-form__contenteditable, [contenteditable='true']"))
                            )
                            message_input.send_keys(formatted_message)
                            time.sleep(random.uniform(0.3, 0.6))
                            send_button = WebDriverWait(self.driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, ".msg-form__send-button"))
                            )
                            ActionChains(self.driver).move_to_element(send_button).click().perform()
                            self.log(f"Sent message to {contact['name']}")
                            time.sleep(random.uniform(1, 2))

                            # Close message window
                            close_button = WebDriverWait(self.driver, 10).until(
                                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[aria-label*='Close your conversation'], button[aria-label*='Dismiss']"))
                            )
                            ActionChains(self.driver).move_to_element(close_button).click().perform()
                            self.log(f"Closed message window for {contact['name']}")
                            time.sleep(random.uniform(0.3, 0.6))

                            successful_sends += 1
                            break  # Exit retry loop on success
                        except Exception as e:
                            self.log(f"Failed to send message to {contact['name']} on retry {retry + 1}: {str(e)}")
                            if retry < max_retries - 1:
                                self.log(f"Retrying {contact['name']} after delay")
                                time.sleep(random.uniform(5, 10))
                            # Take screenshot on final failure
                            if retry == max_retries - 1:
                                try:
                                    screenshot_path = f"screenshot_{contact['name'].replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                                    self.driver.save_screenshot(screenshot_path)
                                    self.log(f"Saved screenshot: {screenshot_path}")
                                except Exception as e:
                                    self.log(f"Failed to save screenshot for {contact['name']}: {str(e)}")
                            continue
                        finally:
                            self.send_progress["value"] = i + 1
                            self.root.update()

                self.log(f"Messages sent successfully to {successful_sends} contacts")
            except Exception as e:
                self.log(f"Error sending messages: {str(e)}")
                # Save debugging artifacts
                try:
                    screenshot_path = f"screenshot_send_messages_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
                    self.driver.save_screenshot(screenshot_path)
                    self.log(f"Saved screenshot: {screenshot_path}")
                except Exception as e:
                    self.log(f"Failed to save screenshot: {str(e)}")
                try:
                    html_snippet = self.driver.page_source[:1000]
                    html_path = f"html_send_messages_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
                    with open(html_path, "w", encoding="utf-8") as f:
                        f.write(html_snippet)
                    self.log(f"Saved page HTML: {html_path}")
                except Exception as e:
                    self.log(f"Failed to save page HTML: {str(e)}")
            finally:
                self.log(f"Messaging completed. Sent {successful_sends} messages, skipped {len(selected_contacts) - successful_sends} contacts")
                self.send_progress["value"] = 0

        # Run in a thread with error handling and timeout
        try:
            self.log("Starting send_messages thread")
            thread = threading.Thread(target=send, daemon=True)
            thread.start()
            thread.join(timeout=300)  # 5-minute timeout
            if thread.is_alive():
                self.log("Messaging thread timed out after 5 minutes")
                self.root.after(0, lambda: messagebox.showerror("Error", "Messaging operation timed out. Please restart the application."))
                # Attempt to clean up
                try:
                    self.driver.quit()
                    self.driver = None
                    self.log("Closed browser due to timeout")
                except Exception as e:
                    self.log(f"Failed to close browser after timeout: {str(e)}")
        except Exception as e:
            self.log(f"Threading error in send_messages: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("Error", f"Threading error: {str(e)}"))

    def __del__(self):
        if self.driver:
            self.driver.quit()

if __name__ == "__main__":
    root = tk.Tk()
    app = LinkedInMessenger(root)
    root.mainloop()
