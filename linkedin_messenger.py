import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import random
import json
import os
import csv
import threading
from datetime import datetime, timedelta
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, NoSuchElementException, ElementNotInteractableException
import re

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
        self.selected_text = scrolledtext.ScrolledText(q3_frame, height=5, width=60, state="disabled")
        self.selected_text.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

        ttk.Label(q3_frame, text="Template:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.template_combo = ttk.Combobox(q3_frame, values=[
            "Hi {first_name}, I noticed your work at {company}. Let's connect!",
            "Hello {first_name}, I'm impressed by your role as {job_title}. Can we chat?",
            "Hi {first_name}, I'm in {industry} too. Let's discuss opportunities!"
        ], width=50)
        self.template_combo.grid(row=2, column=1, padx=5, pady=5)
        self.template_combo.bind("<<ComboboxSelected>>", self.load_template)

        self.message_text = scrolledtext.ScrolledText(q3_frame, height=5, width=60)
        self.message_text.grid(row=3, column=0, columnspan=2, padx=5, pady=5)

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
                options = webdriver.ChromeOptions()
                options.add_argument("--disable-blink-features=AutomationControlled")
                self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
                self.driver.get("https://www.linkedin.com/login")
                time.sleep(2)

                email_field = self.driver.find_element(By.ID, "username")
                email_field.send_keys(email)
                password_field = self.driver.find_element(By.ID, "password")
                password_field.send_keys(password)
                password_field.send_keys(Keys.RETURN)
                time.sleep(5)

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
                self.driver.get("https://www.linkedin.com/mynetwork/invite-connect/connections/")
                time.sleep(5)

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
                        name_elem = elem.find_element(By.CSS_SELECTOR, ".mn-connection-card__name, .connection-card__name, .t-16.t-black.t-bold")
                        name = name_elem.text.strip() if name_elem else ""
                        details_elem = elem.find_element(By.CSS_SELECTOR, ".mn-connection-card__occupation, .connection-card__occupation, .t-14.t-black--light")
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
                if len(self.contacts) < 1000:
                    self.root.after(0, lambda: self.log("Warning: Searched fewer contacts than expected (~1200). LinkedIn may be limiting visibility."))
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
            self.driver.get("https://www.linkedin.com/mynetwork/invite-connect/connections/")
            time.sleep(5)

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
                    name_elem = elem.find_element(By.CSS_SELECTOR, ".mn-connection-card__name, .connection-card__name, .t-16.t-black.t-bold")
                    name = name_elem.text.strip() if name_elem else ""
                    details_elem = elem.find_element(By.CSS_SELECTOR, ".mn-connection-card__occupation, .connection-card__occupation, .t-14.t-black--light")
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
        self.selected_text.configure(state="normal")
        self.selected_text.delete("1.0", tk.END)
        selected_contacts = [c for c in self.contacts if c["selected"]]

        if selected_contacts:
            self.selected_text.insert(tk.END, "Selected Contacts:\n")
            for contact in selected_contacts:
                self.selected_text.insert(tk.END, f"- {contact['name']}\n")
        else:
            self.selected_text.insert(tk.END, "No contacts selected.\n")

        self.selected_text.configure(state="disabled")
        self.selected_text.see(tk.END)

    def check_page_state(self):
        try:
            WebDriverWait(self.driver, 5).until(lambda d: d.execute_script("return document.readyState") == "complete")
            page_title = self.driver.title
            page_url = self.driver.current_url
            if "login" in page_url.lower():
                self.log(f"Session invalid on page '{page_title}' at {page_url}, login required")
                self.root.after(0, lambda: messagebox.showerror("Error", "Session expired. Please log in again."))
                return False
            if self.driver.find_elements(By.CSS_SELECTOR, "[id*='captcha'], [class*='captcha']"):
                self.log(f"CAPTCHA detected on page '{page_title}' at {page_url}. Please resolve manually in the browser and press OK.")
                self.root.after(0, lambda: messagebox.showinfo("CAPTCHA", "Please resolve the CAPTCHA in the browser and press OK to continue."))
                return False
            if self.driver.find_elements(By.XPATH, "//h1[contains(text(), 'restricted')]"):
                self.log(f"Account restricted on page '{page_title}' at {page_url}")
                self.root.after(0, lambda: messagebox.showerror("Error", "Account restricted by LinkedIn. Please resolve and try again."))
                return False
            if self.driver.find_elements(By.XPATH, "//*[contains(text(), 'weekly limit') or contains(text(), 'reached the limit')]"):
                self.log(f"Rate limit detected on page '{page_title}' at {page_url}. Please wait and try again later.")
                self.root.after(0, lambda: messagebox.showwarning("Rate Limit", "LinkedIn has imposed a rate limit. Please wait and try again later."))
                return False
            return True
        except Exception as e:
            self.log(f"Error checking page state on '{self.driver.title}' at {self.driver.current_url}: {str(e)}")
            return False

    def normalize_name(self, name):
        name = re.sub(r'\s+', ' ', name.strip()).lower()
        name = re.sub(r'[^\w\s-]', '', name)
        return name.replace(' ', '-')

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
                self.root.after(0, lambda: self.log("Navigating to connections page"))
                self.driver.get("https://www.linkedin.com/mynetwork/invite-connect/connections/")
                time.sleep(random.uniform(6, 10))
                if not self.check_page_state():
                    self.root.after(0, lambda: self.log("Failed to load connections page. Aborting messaging."))
                    return

                for i, contact in enumerate(selected_contacts):
                    max_retries = 2
                    contact_found = False
                    for attempt in range(max_retries):
                        try:
                            self.root.after(0, lambda: self.log(f"Attempt {attempt + 1}/{max_retries}: Processing {contact['name']} at {self.driver.current_url}"))
                            if not self.check_page_state():
                                self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Page state invalid (CAPTCHA, restriction, or session expired)"))
                                break

                            if "mynetwork/invite-connect/connections" not in self.driver.current_url:
                                self.root.after(0, lambda: self.log(f"Navigating to connections page for {contact['name']}"))
                                self.driver.get("https://www.linkedin.com/mynetwork/invite-connect/connections/")
                                time.sleep(random.uniform(6, 10))
                                if not self.check_page_state():
                                    self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Failed to load connections page"))
                                    break

                            contact_elements = self.driver.find_elements(By.CSS_SELECTOR, ".mn-connection-card, .connection-card, .search-result__wrapper")
                            for elem in contact_elements:
                                try:
                                    name_elem = elem.find_element(By.CSS_SELECTOR, ".mn-connection-card__name, .connection-card__name, .name.actor-name, .t-16.t-black.t-bold")
                                    if name_elem.text.strip() == contact["name"]:
                                        contact["element"] = elem
                                        contact_found = True
                                        self.root.after(0, lambda: self.log(f"Found {contact['name']} on connections page"))
                                        break
                                except StaleElementReferenceException:
                                    continue

                            if contact_found and contact["element"]:
                                contact_element = contact["element"]
                            else:
                                self.root.after(0, lambda: self.log(f"Searching for {contact['name']}"))
                                search_input = None
                                try:
                                    for _ in range(2):
                                        try:
                                            search_input = WebDriverWait(self.driver, 40).until(
                                                EC.presence_of_element_located((By.CSS_SELECTOR, "input[placeholder*='Search'][type='text']"))
                                            )
                                            input_attrs = self.driver.execute_script("return {id: arguments[0].id, class: arguments[0].className, placeholder: arguments[0].placeholder};", search_input)
                                            self.root.after(0, lambda: self.log(f"Located search input for {contact['name']}: id={input_attrs['id']}, class={input_attrs['class']}, placeholder={input_attrs['placeholder']}"))
                                            break
                                        except TimeoutException:
                                            self.root.after(0, lambda: self.log(f"CSS selector failed, trying XPath for search input"))
                                            search_input = WebDriverWait(self.driver, 40).until(
                                                EC.presence_of_element_located((By.XPATH, "//input[contains(@placeholder, 'Search') and @type='text']"))
                                            )
                                            input_attrs = self.driver.execute_script("return {id: arguments[0].id, class: arguments[0].className, placeholder: arguments[0].placeholder};", search_input)
                                            self.root.after(0, lambda: self.log(f"Located search input via XPath for {contact['name']}: id={input_attrs['id']}, class={input_attrs['class']}, placeholder={input_attrs['placeholder']}"))
                                            break
                                    if not search_input:
                                        raise TimeoutException("Search input not found after retries")
                                    search_input.clear()
                                    time.sleep(random.uniform(0.5, 1.0))
                                    for char in contact["name"]:
                                        search_input.send_keys(char)
                                        time.sleep(random.uniform(0.1, 0.2))
                                    time.sleep(random.uniform(2, 3))
                                    search_input.send_keys(Keys.RETURN)
                                    time.sleep(random.uniform(6, 10))

                                    try:
                                        # Check for iframe and switch to it
                                        try:
                                            WebDriverWait(self.driver, 5).until(
                                                EC.frame_to_be_available_and_switch_to_it((By.TAG_NAME, "iframe"))
                                            )
                                            self.root.after(0, lambda: self.log(f"Switched to iframe for {contact['name']}"))
                                        except TimeoutException:
                                            self.root.after(0, lambda: self.log(f"No iframe found for {contact['name']}, continuing in default content"))
                                            self.driver.switch_to.default_content()

                                        # Wait for search results container
                                        WebDriverWait(self.driver, 40).until(
                                            EC.presence_of_element_located((
                                                By.CSS_SELECTOR, 
                                                ".search-results-container, .scaffold-layout__list-container, section[class*='connections'], ul[class*='connections'], div[class*='search-results']"
                                            ))
                                        )
                                        self.root.after(0, lambda: self.log(f"Search results container found for {contact['name']}"))
                                    except TimeoutException:
                                        self.root.after(0, lambda: self.log(f"No search results container found for {contact['name']}"))

                                    contact_elements = self.driver.find_elements(By.CSS_SELECTOR, ".search-result__wrapper, li[class*='search-result'], li.reusable-search__result-container, .mn-connection-card, .connection-card")
                                    found_names = []
                                    for elem in contact_elements:
                                        try:
                                            name_elem = elem.find_element(By.CSS_SELECTOR, ".name.actor-name, .entity-result__title-text, .mn-connection-card__name, .connection-card__name, .t-16.t-black.t-bold")
                                            found_names.append(name_elem.text.strip())
                                        except (NoSuchElementException, StaleElementReferenceException):
                                            continue
                                    self.root.after(0, lambda: self.log(f"Found {len(found_names)} contacts in search results: {', '.join(found_names)}"))

                                    contact_element = None
                                    for elem in contact_elements:
                                        try:
                                            name_elem = elem.find_element(By.CSS_SELECTOR, ".name.actor-name, .entity-result__title-text, .mn-connection-card__name, .connection-card__name, .t-16.t-black.t-bold")
                                            if name_elem.text.strip() == contact["name"]:
                                                contact_element = elem
                                                contact_found = True
                                                self.root.after(0, lambda: self.log(f"Found {contact['name']} in search results"))
                                                break
                                        except StaleElementReferenceException:
                                            continue
                                    
                                    # Switch back to default content
                                    self.driver.switch_to.default_content()
                                    self.root.after(0, lambda: self.log(f"Switched back to default content for {contact['name']}"))

                                    if not contact_element:
                                        self.root.after(0, lambda: self.log(f"No contact element found for {contact['name']} in search results"))
                                        try:
                                            search_area = self.driver.find_element(By.CSS_SELECTOR, ".search-results-container, .scaffold-layout__list-container, section[class*='connections'], ul[class*='connections'], div[class*='search-results']").get_attribute("outerHTML")[:500]
                                            self.root.after(0, lambda: self.log(f"Search area HTML: {search_area}"))
                                        except Exception:
                                            self.root.after(0, lambda: self.log(f"Failed to capture search area HTML at {self.driver.current_url} (Title: {self.driver.title})"))
                                        if attempt < max_retries - 1:
                                            self.root.after(0, lambda: self.log(f"Retrying search for {contact['name']}"))
                                            continue
                                        normalized_name = self.normalize_name(contact["name"])
                                        profile_url = f"https://www.linkedin.com/in/{normalized_name}/"
                                        self.root.after(0, lambda: self.log(f"Attempting profile page for {contact['name']}: {profile_url}"))
                                        self.driver.get(profile_url)
                                        time.sleep(random.uniform(6, 10))
                                        if not self.check_page_state():
                                            self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Failed to load profile page"))
                                            break
                                        try:
                                            name_elem = WebDriverWait(self.driver, 40).until(
                                                EC.presence_of_element_located((By.CSS_SELECTOR, "h1.text-heading-xlarge"))
                                            )
                                            if name_elem.text.strip() == contact["name"]:
                                                contact_element = self.driver.find_element(By.CSS_SELECTOR, "main")
                                                contact_found = True
                                                self.root.after(0, lambda: self.log(f"Verified {contact['name']} on profile page"))
                                            else:
                                                self.root.after(0, lambda: self.log(f"Profile page name mismatch for {contact['name']}: Found '{name_elem.text.strip()}'"))
                                                break
                                        except Exception as e:
                                            self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Could not verify profile page: {str(e)}"))
                                            break
                                    if not contact_found:
                                        self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Could not find contact on LinkedIn after search"))
                                        break
                                except TimeoutException:
                                    self.root.after(0, lambda: self.log(f"Attempt {attempt + 1}/{max_retries}: Search input not found for {contact['name']}"))
                                    if attempt < max_retries - 1:
                                        self.driver.refresh()
                                        time.sleep(random.uniform(2, 4))
                                        continue
                                    self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Search input not found after retries"))
                                    break
                                except ElementNotInteractableException:
                                    self.root.after(0, lambda: self.log(f"Attempt {attempt + 1}/{max_retries}: Search input not interactable for {contact['name']}"))
                                    if attempt < max_retries - 1:
                                        self.driver.refresh()
                                        time.sleep(random.uniform(2, 4))
                                        continue
                                    self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Search input not interactable after retries"))
                                    break
                                except Exception as e:
                                    error_msg = str(e)
                                    self.root.after(0, lambda: self.log(f"Attempt {attempt + 1}/{max_retries}: Failed to search for {contact['name']}: {error_msg}"))
                                    if attempt < max_retries - 1:
                                        self.driver.refresh()
                                        time.sleep(random.uniform(2, 4))
                                        continue
                                    self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Failed to search: {error_msg}"))
                                    break
                            if not contact_found:
                                self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Could not find contact after attempts"))
                                normalized_name = self.normalize_name(contact["name"])
                                profile_url = f"https://www.linkedin.com/in/{normalized_name}/"
                                self.root.after(0, lambda: self.log(f"Forcing profile page for {contact['name']}: {profile_url}"))
                                self.driver.get(profile_url)
                                time.sleep(random.uniform(6, 10))
                                if not self.check_page_state():
                                    self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Failed to load profile page"))
                                    break
                                try:
                                    name_elem = WebDriverWait(self.driver, 40).until(
                                        EC.presence_of_element_located((By.CSS_SELECTOR, "h1.text-heading-xlarge"))
                                    )
                                    if name_elem.text.strip() == contact["name"]:
                                        contact_element = self.driver.find_element(By.CSS_SELECTOR, "main")
                                        contact_found = True
                                        self.root.after(0, lambda: self.log(f"Verified {contact['name']} on profile page"))
                                    else:
                                        self.root.after(0, lambda: self.log(f"Profile page name mismatch for {contact['name']}: Found '{name_elem.text.strip()}'"))
                                        break
                                except Exception as e:
                                    self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Could not verify profile page: {str(e)}"))
                                    break
                            if not contact_found:
                                self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Could not find contact on profile page"))
                                break

                            try:
                                if not self.check_page_state():
                                    self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Page state invalid before clicking message button"))
                                    break
                                self.root.after(0, lambda: self.log(f"Locating message button for {contact['name']}"))
                                message_button = WebDriverWait(self.driver, 40).until(
                                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Message') or contains(@aria-label, 'Message') or contains(@class, 'msg') or @data-control-name='message']"))
                                )
                                self.driver.execute_script("arguments[0].scrollIntoView(true);", message_button)
                                time.sleep(random.uniform(0.5, 1.0))
                                message_button.click()
                                self.root.after(0, lambda: self.log(f"Clicked message button for {contact['name']}"))
                                time.sleep(random.uniform(2, 3))
                            except Exception as e:
                                error_msg = str(e)
                                self.root.after(0, lambda: self.log(f"Attempt {attempt + 1}/{max_retries}: Failed to find or click message button for {contact['name']}: {error_msg}"))
                                if attempt < max_retries - 1:
                                    self.driver.refresh()
                                    time.sleep(random.uniform(2, 4))
                                    continue
                                self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Failed to find or click message button: {error_msg}"))
                                break

                            first_name = contact["name"].split()[0]
                            formatted_message = message.format(
                                first_name=first_name,
                                name=contact["name"],
                                job_title=contact["job_title"],
                                company=contact["company"],
                                industry=contact["industry"]
                            )
                            try:
                                if not self.check_page_state():
                                    self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Page state invalid before sending message"))
                                    break
                                self.root.after(0, lambda: self.log(f"Filling message input for {contact['name']}"))
                                message_input = WebDriverWait(self.driver, 40).until(
                                    EC.presence_of_element_located((By.CSS_SELECTOR, ".msg-form__contenteditable"))
                                )
                                message_input.send_keys(formatted_message)
                                self.root.after(0, lambda: self.log(f"Message input filled for {contact['name']}"))
                                time.sleep(random.uniform(0.5, 1.5))
                                send_button = WebDriverWait(self.driver, 40).until(
                                    EC.element_to_be_clickable((By.CSS_SELECTOR, ".msg-form__send-button"))
                                )
                                send_button.click()
                                self.root.after(0, lambda: self.log(f"Clicked send button for {contact['name']}"))
                                time.sleep(random.uniform(2, 3))
                            except Exception as e:
                                error_msg = str(e)
                                self.root.after(0, lambda: self.log(f"Attempt {attempt + 1}/{max_retries}: Failed to send message to {contact['name']}: {error_msg}"))
                                if attempt < max_retries - 1:
                                    self.driver.refresh()
                                    time.sleep(random.uniform(2, 4))
                                    continue
                                self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Failed to send message: {error_msg}"))
                                break

                            try:
                                if not self.check_page_state():
                                    self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Page state invalid before closing message window"))
                                    break
                                self.root.after(0, lambda: self.log(f"Closing message window for {contact['name']}"))
                                close_button = WebDriverWait(self.driver, 40).until(
                                    EC.element_to_be_clickable((By.CSS_SELECTOR, "button[aria-label*='Close your conversation'], button[aria-label*='Dismiss']"))
                                )
                                close_button.click()
                                self.root.after(0, lambda: self.log(f"Closed message window for {contact['name']}"))
                                time.sleep(random.uniform(0.5, 1.5))
                            except Exception as e:
                                error_msg = str(e)
                                self.root.after(0, lambda: self.log(f"Attempt {attempt + 1}/{max_retries}: Failed to close message window for {contact['name']}: {error_msg}"))
                                if attempt < max_retries - 1:
                                    self.driver.refresh()
                                    time.sleep(random.uniform(2, 4))
                                    continue
                                self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Failed to close message window: {error_msg}"))
                                break

                            self.root.after(0, lambda: self.log(f"Sent message to {contact['name']}"))
                            successful_sends += 1
                            break
                        except Exception as e:
                            error_msg = str(e)
                            self.root.after(0, lambda: self.log(f"Attempt {attempt + 1}/{max_retries}: Unexpected error for {contact['name']} at {self.driver.current_url}: {error_msg}"))
                            if attempt < max_retries - 1:
                                self.driver.refresh()
                                time.sleep(random.uniform(2, 4))
                                continue
                            self.root.after(0, lambda: self.log(f"Skipping {contact['name']}: Unexpected error after retries: {error_msg}"))
                            break
                    self.root.after(0, lambda: self.send_progress.configure(value=i + 1))
                    self.root.update()
                self.root.after(0, lambda: self.log(f"Messages sent successfully to {successful_sends} contacts"))
            except Exception as e:
                self.root.after(0, lambda: self.log(f"Error sending messages: {str(e)}"))
            finally:
                self.root.after(0, lambda: self.log(f"Messaging process completed. Sent {successful_sends} messages, skipped {len(selected_contacts) - successful_sends} contacts"))
                self.root.after(0, lambda: self.send_progress.configure(value=0))

        threading.Thread(target=send, daemon=True).start()

    def __del__(self):
        if self.driver:
            self.driver.quit()

if __name__ == "__main__":
    root = tk.Tk()
    app = LinkedInMessenger(root)
    root.mainloop()
