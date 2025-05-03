import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import re
import random
from datetime import datetime

class LinkedInMessenger:
    def __init__(self, root):
        self.root = root
        self.root.title("LinkedIn Messenger")
        self.root.geometry("800x600")
        self.driver = None
        self.contacts = []
        self.setup_gui()

    def setup_gui(self):
        # Login Frame
        login_frame = ttk.LabelFrame(self.root, text="LinkedIn Login", padding=10)
        login_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(login_frame, text="Email:").grid(row=0, column=0, padx=5, pady=5)
        self.email_entry = ttk.Entry(login_frame, width=40)
        self.email_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(login_frame, text="Password:").grid(row=1, column=0, padx=5, pady=5)
        self.password_entry = ttk.Entry(login_frame, width=40, show="*")
        self.password_entry.grid(row=1, column=1, padx=5, pady=5)

        ttk.Button(login_frame, text="Login", command=self.login_linkedin).grid(row=2, column=0, columnspan=2, pady=10)

        # Filter Frame
        filter_frame = ttk.LabelFrame(self.root, text="Filter Contacts", padding=10)
        filter_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(filter_frame, text="Name:").grid(row=0, column=0, padx=5, pady=5)
        self.name_filter = ttk.Entry(filter_frame)
        self.name_filter.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(filter_frame, text="Job Title:").grid(row=0, column=2, padx=5, pady=5)
        self.job_filter = ttk.Entry(filter_frame)
        self.job_filter.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(filter_frame, text="Industry:").grid(row=1, column=0, padx=5, pady=5)
        self.industry_filter = ttk.Entry(filter_frame)
        self.industry_filter.grid(row=1, column=1, padx=5, pady=5)

        ttk.Label(filter_frame, text="Company:").grid(row=1, column=2, padx=5, pady=5)
        self.company_filter = ttk.Entry(filter_frame)
        self.company_filter.grid(row=1, column=3, padx=5, pady=5)

        ttk.Button(filter_frame, text="Fetch Contacts", command=self.fetch_contacts).grid(row=2, column=0, columnspan=4, pady=10)

        # Contacts Frame
        contacts_frame = ttk.LabelFrame(self.root, text="Select Contacts", padding=10)
        contacts_frame.pack(fill="both", expand=True, padx=5, pady=5)

        self.contacts_tree = ttk.Treeview(contacts_frame, columns=("Name", "Job Title", "Company", "Industry"), show="headings")
        self.contacts_tree.heading("Name", text="Name")
        self.contacts_tree.heading("Job Title", text="Job Title")
        self.contacts_tree.heading("Company", text="Company")
        self.contacts_tree.heading("Industry", text="Industry")
        self.contacts_tree.pack(fill="both", expand=True)

        ttk.Button(contacts_frame, text="Select All", command=self.select_all_contacts).pack(side="left", padx=5)
        ttk.Button(contacts_frame, text="Deselect All", command=self.deselect_all_contacts).pack(side="left", padx=5)

        # Message Frame
        message_frame = ttk.LabelFrame(self.root, text="Message", padding=10)
        message_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(message_frame, text="Template:").grid(row=0, column=0, padx=5, pady=5)
        self.template_combo = ttk.Combobox(message_frame, values=[
            "Hi {first_name}, I noticed your work at {company}. Let's connect!",
            "Hello {first_name}, I'm impressed by your role as {job_title}. Can we chat?",
            "Hi {first_name}, I'm in {industry} too. Let's discuss opportunities!"
        ], width=50)
        self.template_combo.grid(row=0, column=1, padx=5, pady=5)
        self.template_combo.bind("<<ComboboxSelected>>", self.load_template)

        self.message_text = scrolledtext.ScrolledText(message_frame, height=5, width=60)
        self.message_text.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

        ttk.Button(message_frame, text="Preview", command=self.preview_message).grid(row=2, column=0, padx=5, pady=5)
        ttk.Button(message_frame, text="Send Messages", command=self.send_messages).grid(row=2, column=1, padx=5, pady=5)

        # Progress and Log
        self.progress = ttk.Progressbar(self.root, mode="determinate")
        self.progress.pack(fill="x", padx=5, pady=5)

        self.log_text = scrolledtext.ScrolledText(self.root, height=5, state="disabled")
        self.log_text.pack(fill="x", padx=5, pady=5)

    def log(self, message):
        self.log_text.configure(state="normal")
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}\n")
        self.log_text.configure(state="disabled")
        self.log_text.see(tk.END)

    def login_linkedin(self):
        email = self.email_entry.get()
        password = self.password_entry.get()
        if not email or not password:
            messagebox.showerror("Error", "Please enter email and password")
            return

        try:
            self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
            self.driver.get("https://www.linkedin.com/login")
            time.sleep(2)

            email_field = self.driver.find_element(By.ID, "username")
            email_field.send_keys(email)
            password_field = self.driver.find_element(By.ID, "password")
            password_field.send_keys(password)
            password_field.send_keys(Keys.RETURN)
            time.sleep(5)

            if "login" in self.driver.current_url:
                messagebox.showerror("Error", "Login failed. Check credentials.")
                self.driver.quit()
                self.driver = None
                return

            self.log("Logged in successfully")
            messagebox.showinfo("Success", "Logged in to LinkedIn")
        except Exception as e:
            self.log(f"Login error: {str(e)}")
            messagebox.showerror("Error", f"Login failed: {str(e)}")
            if self.driver:
                self.driver.quit()
                self.driver = None

    def fetch_contacts(self):
        if not self.driver:
            messagebox.showerror("Error", "Please log in first")
            return

        try:
            self.contacts_tree.delete(*self.contacts_tree.get_children())
            self.contacts = []
            self.driver.get("https://www.linkedin.com/mynetwork/invite-connect/connections/")
            time.sleep(3)

            # Scroll to load more contacts
            last_height = self.driver.execute_script("return document.body.scrollHeight")
            for _ in range(3):  # Scroll 3 times
                self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2)
                new_height = self.driver.execute_script("return document.body.scrollHeight")
                if new_height == last_height:
                    break
                last_height = new_height

            contact_elements = self.driver.find_elements(By.CSS_SELECTOR, ".mn-connection-card")
            for elem in contact_elements:
                try:
                    name = elem.find_element(By.CSS_SELECTOR, ".mn-connection-card__name").text
                    details = elem.find_element(By.CSS_SELECTOR, ".mn-connection-card__occupation").text
                    # Parse details for job title, company, industry (approximate)
                    job_title = details.split(" at ")[0].strip() if " at " in details else ""
                    company = details.split(" at ")[1].strip() if " at " in details else ""
                    industry = ""  # LinkedIn doesn't always show industry; leave blank or infer
                    contact = {
                        "name": name,
                        "job_title": job_title,
                        "company": company,
                        "industry": industry,
                        "element": elem
                    }
                    self.contacts.append(contact)
                except:
                    continue

            # Apply filters
            name_filter = self.name_filter.get().lower()
            job_filter = self.job_filter.get().lower()
            industry_filter = self.industry_filter.get().lower()
            company_filter = self.company_filter.get().lower()

            filtered_contacts = self.contacts
            if name_filter:
                filtered_contacts = [c for c in filtered_contacts if name_filter in c["name"].lower()]
            if job_filter:
                filtered_contacts = [c for c in filtered_contacts if job_filter in c["job_title"].lower()]
            if industry_filter:
                filtered_contacts = [c for c in filtered_contacts if industry_filter in c["industry"].lower()]
            if company_filter:
                filtered_contacts = [c for c in filtered_contacts if company_filter in c["company"].lower()]

            for contact in filtered_contacts:
                self.contacts_tree.insert("", "end", values=(
                    contact["name"],
                    contact["job_title"],
                    contact["company"],
                    contact["industry"]
                ))

            self.log(f"Fetched {len(filtered_contacts)} contacts")
        except Exception as e:
            self.log(f"Error fetching contacts: {str(e)}")
            messagebox.showerror("Error", f"Failed to fetch contacts: {str(e)}")

    def select_all_contacts(self):
        for item in self.contacts_tree.get_children():
            self.contacts_tree.selection_add(item)

    def deselect_all_contacts(self):
        self.contacts_tree.selection_remove(self.contacts_tree.selection())

    def load_template(self, event):
        self.message_text.delete("1.0", tk.END)
        self.message_text.insert("1.0", self.template_combo.get())

    def preview_message(self):
        selected = self.contacts_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "No contacts selected")
            return

        message = self.message_text.get("1.0", tk.END).strip()
        if not message:
            messagebox.showwarning("Warning", "No message entered")
            return

        preview = ""
        for item in selected[:3]:  # Show preview for up to 3 contacts
            values = self.contacts_tree.item(item)["values"]
            name = values[0]
            first_name = name.split()[0]
            job_title = values[1]
            company = values[2]
            industry = values[3]
            formatted_message = message.format(
                first_name=first_name,
                name=name,
                job_title=job_title,
                company=company,
                industry=industry
            )
            preview += f"To {name}:\n{formatted_message}\n\n"

        if len(selected) > 3:
            preview += f"...and {len(selected) - 3} more contacts"

        messagebox.showinfo("Message Preview", preview)

    def send_messages(self):
        selected = self.contacts_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "No contacts selected")
            return

        message = self.message_text.get("1.0", tk.END).strip()
        if not message:
            messagebox.showwarning("Warning", "No message entered")
            return

        if not messagebox.askyesno("Confirm", f"Send message to {len(selected)} contacts?"):
            return

        self.progress["maximum"] = len(selected)
        self.progress["value"] = 0

        try:
            for i, item in enumerate(selected):
                values = self.contacts_tree.item(item)["values"]
                name = values[0]
                first_name = name.split()[0]
                job_title = values[1]
                company = values[2]
                industry = values[3]

                # Find corresponding contact
                contact = next(c for c in self.contacts if c["name"] == name and c["job_title"] == job_title)
                element = contact["element"]

                # Click message button
                message_button = element.find_element(By.CSS_SELECTOR, "button[aria-label*='Message']")
                message_button.click()
                time.sleep(2)

                # Format and send message
                formatted_message = message.format(
                    first_name=first_name,
                    name=name,
                    job_title=job_title,
                    company=company,
                    industry=industry
                )
                message_input = self.driver.find_element(By.CSS_SELECTOR, ".msg-form__contenteditable")
                message_input.send_keys(formatted_message)
                send_button = self.driver.find_element(By.CSS_SELECTOR, ".msg-form__send-button")
                send_button.click()
                time.sleep(random.uniform(3, 5))  # Random delay to mimic human behavior

                # Close message window
                close_button = self.driver.find_element(By.CSS_SELECTOR, "button[aria-label*='Dismiss']")
                close_button.click()
                time.sleep(1)

                self.log(f"Sent message to {name}")
                self.progress["value"] += 1
                self.root.update()

            messagebox.showinfo("Success", "Messages sent successfully")
        except Exception as e:
            self.log(f"Error sending messages: {str(e)}")
            messagebox.showerror("Error", f"Failed to send messages: {str(e)}")
        finally:
            self.progress["value"] = 0

    def __del__(self):
        if self.driver:
            self.driver.quit()

if __name__ == "__main__":
    root = tk.Tk()
    app = LinkedInMessenger(root)
    root.mainloop()