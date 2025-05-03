import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import re
import random
from datetime import datetime
from tkinter import TclError

class LinkedInMessenger:
    def __init__(self, root):
        self.root = root
        self.root.title("LinkedIn Messenger")
        self.root.geometry("1000x700")
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

        filters = [
            ("Title", "title_filter"), ("First Name", "first_name_filter"), ("Last Name", "last_name_filter"),
            ("Email", "email_filter"), ("Phone", "phone_filter"), ("Job Title", "job_filter"),
            ("Industry", "industry_filter"), ("Company", "company_filter")
        ]

        for i, (label, attr) in enumerate(filters):
            ttk.Label(filter_frame, text=f"{label}:").grid(row=i//4, column=(i%4)*2, padx=5, pady=5)
            entry = ttk.Entry(filter_frame)
            entry.grid(row=i//4, column=(i%4)*2+1, padx=5, pady=5)
            setattr(self, attr, entry)
            # Autocomplete binding
            entry.bind("<KeyRelease>", lambda e, f=attr: self.autocomplete(f))

        ttk.Button(filter_frame, text="Fetch Contacts", command=self.fetch_contacts).grid(row=2, column=0, columnspan=8, pady=10)

        # Contacts Frame
        contacts_frame = ttk.LabelFrame(self.root, text="Select Contacts", padding=10)
        contacts_frame.pack(fill="both", expand=True, padx=5, pady=5)

        columns = ("Select", "Title", "First Name", "Last Name", "Email", "Phone", "Job Title", "Company", "Industry")
        self.contacts_tree = ttk.Treeview(contacts_frame, columns=columns, show="headings")
        for col in columns:
            self.contacts_tree.heading(col, text=col, command=lambda c=col: self.sort_treeview(c))
            self.contacts_tree.column(col, width=100, anchor="w")
        self.contacts_tree.column("Select", width=50)
        self.contacts_tree.pack(fill="both", expand=True)

        # Bind checkbox toggle
        self.contacts_tree.bind("<Button-1>", self.toggle_checkbox)

        button_frame = ttk.Frame(contacts_frame)
        button_frame.pack(fill="x", pady=5)
        ttk.Button(button_frame, text="Select All", command=self.select_all_contacts).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Deselect All", command=self.deselect_all_contacts).pack(side="left", padx=5)

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
            for _ in range(3):
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
                    first_name = name.split()[0] if name else ""
                    last_name = " ".join(name.split()[1:]) if len(name.split()) > 1 else ""
                    details = elem.find_element(By.CSS_SELECTOR, ".mn-connection-card__occupation").text
                    job_title = details.split(" at ")[0].strip() if " at " in details else ""
                    company = details.split(" at ")[1].strip() if " at " in details else ""
                    # Placeholder for fields not always available
                    title = ""  # E.g., Mr., Ms., Dr. (not typically in LinkedIn UI)
                    email = ""  # Requires profile access
                    phone = ""  # Requires profile access
                    industry = ""  # Infer or leave blank
                    contact = {
                        "title": title,
                        "first_name": first_name,
                        "last_name": last_name,
                        "email": email,
                        "phone": phone,
                        "job_title": job_title,
                        "company": company,
                        "industry": industry,
                        "element": elem,
                        "selected": False
                    }
                    self.contacts.append(contact)
                except:
                    continue

            # Apply filters
            filters = {
                "title": self.title_filter.get().lower(),
                "first_name": self.first_name_filter.get().lower(),
                "last_name": self.last_name_filter.get().lower(),
                "email": self.email_filter.get().lower(),
                "phone": self.phone_filter.get().lower(),
                "job_title": self.job_filter.get().lower(),
                "industry": self.industry_filter.get().lower(),
                "company": self.company_filter.get().lower()
            }

            filtered_contacts = self.contacts
            for key, value in filters.items():
                if value:
                    filtered_contacts = [c for c in filtered_contacts if value in c[key].lower()]

            for contact in filtered_contacts:
                self.contacts_tree.insert("", "end", values=(
                    "☐",  # Checkbox (unselected)
                    contact["title"],
                    contact["first_name"],
                    contact["last_name"],
                    contact["email"],
                    contact["phone"],
                    contact["job_title"],
                    contact["company"],
                    contact["industry"]
                ))

            self.log(f"Fetched {len(filtered_contacts)} contacts")
        except Exception as e:
            self.log(f"Error fetching contacts: {str(e)}")
            messagebox.showerror("Error", f"Failed to fetch contacts: {str(e)}")

    def autocomplete(self, filter_attr):
        entry = getattr(self, filter_attr)
        current_text = entry.get().lower()
        if not current_text:
            return

        suggestions = set()
        for contact in self.contacts:
            value = contact[filter_attr.replace("_filter", "")].lower()
            if value and current_text in value:
                suggestions.add(contact[filter_attr.replace("_filter", "")])

        if suggestions:
            suggestion_text = f"Suggestions: {', '.join(list(suggestions)[:5])}"
            self.log(suggestion_text)
        else:
            self.log("No matching suggestions")

    def sort_treeview(self, col):
        items = [(self.contacts_tree.set(k, col), k) for k in self.contacts_tree.get_children()]
        items.sort()
        for index, (val, k) in enumerate(items):
            self.contacts_tree.move(k, "", index)

    def toggle_checkbox(self, event):
        try:
            item = self.contacts_tree.identify_row(event.y)
            if not item:
                return
            column = self.contacts_tree.identify_column(event.x)
            if column != "#1":  # Only toggle on Select column
                return

            values = list(self.contacts_tree.item(item)["values"])
            current = values[0]
            new_state = "☑" if current == "☐" else "☐"
            values[0] = new_state
            self.contacts_tree.item(item, values=values)

            # Update contact's selected state
            for contact in self.contacts:
                if (contact["first_name"] == values[2] and
                    contact["last_name"] == values[3] and
                    contact["job_title"] == values[6]):
                    contact["selected"] = (new_state == "☑")
        except TclError:
            pass

    def select_all_contacts(self):
        for item in self.contacts_tree.get_children():
            values = list(self.contacts_tree.item(item)["values"])
            values[0] = "☑"
            self.contacts_tree.item(item, values=values)
            for contact in self.contacts:
                if (contact["first_name"] == values[2] and
                    contact["last_name"] == values[3] and
                    contact["job_title"] == values[6]):
                    contact["selected"] = True

    def deselect_all_contacts(self):
        for item in self.contacts_tree.get_children():
            values = list(self.contacts_tree.item(item)["values"])
            values[0] = "☐"
            self.contacts_tree.item(item, values=values)
            for contact in self.contacts:
                if (contact["first_name"] == values[2] and
                    contact["last_name"] == values[3] and
                    contact["job_title"] == values[6]):
                    contact["selected"] = False

    def load_template(self, event):
        self.message_text.delete("1.0", tk.END)
        self.message_text.insert("1.0", self.template_combo.get())

    def preview_message(self):
        selected_contacts = [c for c in self.contacts if c["selected"]]
        if not selected_contacts:
            messagebox.showwarning("Warning", "No contacts selected")
            return

        message = self.message_text.get("1.0", tk.END).strip()
        if not message:
            messagebox.showwarning("Warning", "No message entered")
            return

        preview = ""
        for contact in selected_contacts[:3]:
            formatted_message = message.format(
                title=contact["title"],
                first_name=contact["first_name"],
                last_name=contact["last_name"],
                email=contact["email"],
                phone=contact["phone"],
                job_title=contact["job_title"],
                company=contact["company"],
                industry=contact["industry"]
            )
            preview += f"To {contact['first_name']} {contact['last_name']}:\n{formatted_message}\n\n"

        if len(selected_contacts) > 3:
            preview += f"...and {len(selected_contacts) - 3} more contacts"

        messagebox.showinfo("Message Preview", preview)

    def send_messages(self):
        selected_contacts = [c for c in self.contacts if c["selected"]]
        if not selected_contacts:
            messagebox.showwarning("Warning", "No contacts selected")
            return

        message = self.message_text.get("1.0", tk.END).strip()
        if not message:
            messagebox.showwarning("Warning", "No message entered")
            return

        if not messagebox.askyesno("Confirm", f"Send message to {len(selected_contacts)} contacts?"):
            return

        self.progress["maximum"] = len(selected_contacts)
        self.progress["value"] = 0

        try:
            for contact in selected_contacts:
                element = contact["element"]
                message_button = element.find_element(By.CSS_SELECTOR, "button[aria-label*='Message']")
                message_button.click()
                time.sleep(2)

                formatted_message = message.format(
                    title=contact["title"],
                    first_name=contact["first_name"],
                    last_name=contact["last_name"],
                    email=contact["email"],
                    phone=contact["phone"],
                    job_title=contact["job_title"],
                    company=contact["company"],
                    industry=contact["industry"]
                )
                message_input = self.driver.find_element(By.CSS_SELECTOR, ".msg-form__contenteditable")
                message_input.send_keys(formatted_message)
                send_button = self.driver.find_element(By.CSS_SELECTOR, ".msg-form__send-button")
                send_button.click()
                time.sleep(random.uniform(3, 5))

                close_button = self.driver.find_element(By.CSS_SELECTOR, "button[aria-label*='Dismiss']")
                close_button.click()
                time.sleep(1)

                self.log(f"Sent message to {contact['first_name']} {contact['last_name']}")
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
