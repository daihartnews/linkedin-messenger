import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
import random
from datetime import datetime

class LinkedInMessenger:
    def __init__(self, root):
        self.root = root
        self.root.title("LinkedIn Messenger")
        self.root.geometry("1000x700")
        self.driver = None
        self.contacts = []
        self.sort_column = None
        self.sort_reverse = False
        self.setup_gui()

    def setup_gui(self):
        # Main layout: Split into four quadrants using PanedWindow
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(fill="both", expand=True)

        left_pane = ttk.PanedWindow(main_pane, orient=tk.VERTICAL)
        right_pane = ttk.PanedWindow(main_pane, orient=tk.VERTICAL)
        main_pane.add(left_pane, weight=1)
        main_pane.add(right_pane, weight=1)

        # 1st Quadrant: Login and Filters
        q1_frame = ttk.LabelFrame(left_pane, text="Login & Filters", padding=10)
        left_pane.add(q1_frame, weight=1)

        # Login Section
        ttk.Label(q1_frame, text="Email:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.email_entry = ttk.Entry(q1_frame, width=30)
        self.email_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(q1_frame, text="Password:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.password_entry = ttk.Entry(q1_frame, width=30, show="*")
        self.password_entry.grid(row=1, column=1, padx=5, pady=5)

        ttk.Button(q1_frame, text="Login", command=self.login_linkedin).grid(row=2, column=0, columnspan=2, pady=10)

        # Filter Section
        ttk.Label(q1_frame, text="Name:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.name_filter = ttk.Entry(q1_frame)
        self.name_filter.grid(row=3, column=1, padx=5, pady=5)

        ttk.Label(q1_frame, text="Job Title:").grid(row=4, column=0, padx=5, pady=5, sticky="e")
        self.job_filter = ttk.Entry(q1_frame)
        self.job_filter.grid(row=4, column=1, padx=5, pady=5)

        ttk.Label(q1_frame, text="Company:").grid(row=5, column=0, padx=5, pady=5, sticky="e")
        self.company_filter = ttk.Entry(q1_frame)
        self.company_filter.grid(row=5, column=1, padx=5, pady=5)

        ttk.Label(q1_frame, text="Industry:").grid(row=6, column=0, padx=5, pady=5, sticky="e")
        self.industry_filter = ttk.Entry(q1_frame)
        self.industry_filter.grid(row=6, column=1, padx=5, pady=5)

        ttk.Button(q1_frame, text="Fetch Contacts", command=self.fetch_contacts).grid(row=7, column=0, columnspan=2, pady=10)

        # 2nd Quadrant: Contacts Table
        q2_frame = ttk.LabelFrame(left_pane, text="Contacts", padding=10)
        left_pane.add(q2_frame, weight=2)

        self.contacts_tree = ttk.Treeview(q2_frame, columns=("Select", "Name", "Job Title", "Company", "Industry"), show="headings")
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
        self.contacts_tree.pack(fill="both", expand=True)

        # Checkbox handling
        self.contacts_tree.bind("<Button-1>", self.handle_tree_click)

        # Select/Deselect Buttons
        button_frame = ttk.Frame(q2_frame)
        button_frame.pack(fill="x", pady=5)
        ttk.Button(button_frame, text="Select All", command=self.select_all_contacts).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Deselect All", command=self.deselect_all_contacts).pack(side="left", padx=5)

        # 3rd Quadrant: Message Composition
        q3_frame = ttk.LabelFrame(right_pane, text="Message", padding=10)
        right_pane.add(q3_frame, weight=1)

        ttk.Label(q3_frame, text="Template:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.template_combo = ttk.Combobox(q3_frame, values=[
            "Hi {first_name}, I noticed your work at {company}. Let's connect!",
            "Hello {first_name}, I'm impressed by your role as {job_title}. Can we chat?",
            "Hi {first_name}, I'm in {industry} too. Let's discuss opportunities!"
        ], width=50)
        self.template_combo.grid(row=0, column=1, padx=5, pady=5)
        self.template_combo.bind("<<ComboboxSelected>>", self.load_template)

        self.message_text = scrolledtext.ScrolledText(q3_frame, height=5, width=60)
        self.message_text.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

        # 4th Quadrant: Preview, Progress, and Logs
        q4_frame = ttk.LabelFrame(right_pane, text="Preview & Status", padding=10)
        right_pane.add(q4_frame, weight=2)

        ttk.Button(q4_frame, text="Preview Messages", command=self.preview_message).pack(fill="x", pady=5)
        ttk.Button(q4_frame, text="Send Messages", command=self.send_messages).pack(fill="x", pady=5)

        self.progress = ttk.Progressbar(q4_frame, mode="determinate")
        self.progress.pack(fill="x", pady=5)

        self.log_text = scrolledtext.ScrolledText(q4_frame, height=10, state="disabled")
        self.log_text.pack(fill="both", expand=True, pady=5)

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
                    details = elem.find_element(By.CSS_SELECTOR, ".mn-connection-card__occupation").text
                    job_title = details.split(" at ")[0].strip() if " at " in details else ""
                    company = details.split(" at ")[1].strip() if " at " in details else ""
                    industry = ""  # LinkedIn doesn't always show industry
                    contact = {
                        "name": name,
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
            name_filter = self.name_filter.get().lower()
            job_filter = self.job_filter.get().lower()
            company_filter = self.company_filter.get().lower()
            industry_filter = self.industry_filter.get().lower()

            filtered_contacts = self.contacts
            if name_filter:
                filtered_contacts = [c for c in filtered_contacts if name_filter in c["name"].lower()]
            if job_filter:
                filtered_contacts = [c for c in filtered_contacts if job_filter in c["job_title"].lower()]
            if company_filter:
                filtered_contacts = [c for c in filtered_contacts if company_filter in c["company"].lower()]
            if industry_filter:
                filtered_contacts = [c for c in filtered_contacts if industry_filter in c["industry"].lower()]

            # Display contacts
            for contact in filtered_contacts:
                self.contacts_tree.insert("", "end", iid=contact["name"], values=(
                    "☐",  # Unchecked box
                    contact["name"],
                    contact["job_title"],
                    contact["company"],
                    contact["industry"]
                ))

            self.log(f"Fetched {len(filtered_contacts)} contacts")
        except Exception as e:
            self.log(f"Error fetching contacts: {str(e)}")
            messagebox.showerror("Error", f"Failed to fetch contacts: {str(e)}")

    def handle_tree_click(self, event):
        # Handle checkbox clicks in the "Select" column
        item = self.contacts_tree.identify_row(event.y)
        column = self.contacts_tree.identify_column(event.x)
        if item and column == "#1":  # Select column
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

    def sort_contacts(self, column):
        # Toggle sort direction if same column, else reset
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_reverse = False
            self.sort_column = column

        # Map column names to contact dictionary keys
        col_map = {
            "Name": "name",
            "Job Title": "job_title",
            "Company": "company",
            "Industry": "industry"
        }
        key = col_map[column]

        # Sort contacts
        sorted_contacts = sorted(self.contacts, key=lambda x: x[key].lower(), reverse=self.sort_reverse)

        # Update Treeview
        self.contacts_tree.delete(*self.contacts_tree.get_children())
        for contact in sorted_contacts:
            self.contacts_tree.insert("", "end", iid=contact["name"], values=(
                "☑" if contact["selected"] else "☐",
                contact["name"],
                contact["job_title"],
                contact["company"],
                contact["industry"]
            ))

    def select_all_contacts(self):
        for contact in self.contacts:
            contact["selected"] = True
        for item in self.contacts_tree.get_children():
            values = self.contacts_tree.item(item)["values"]
            self.contacts_tree.item(item, values=("☑", *values[1:]))

    def deselect_all_contacts(self):
        for contact in self.contacts:
            contact["selected"] = False
        for item in self.contacts_tree.get_children():
            values = self.contacts_tree.item(item)["values"]
            self.contacts_tree.item(item, values=("☐", *values[1:]))

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
        for contact in selected_contacts[:3]:  # Show up to 3 contacts
            first_name = contact["name"].split()[0]
            formatted_message = message.format(
                first_name=first_name,
                name=contact["name"],
                job_title=contact["job_title"],
                company=contact["company"],
                industry=contact["industry"]
            )
            preview += f"To {contact['name']}:\n{formatted_message}\n\n"

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

                first_name = contact["name"].split()[0]
                formatted_message = message.format(
                    first_name=first_name,
                    name=contact["name"],
                    job_title=contact["job_title"],
                    company=contact["company"],
                    industry=contact["industry"]
                )
                message_input = self.driver.find_element(By.CSS_SELECTOR, ".msg-form__contenteditable")
                message_input.send_keys(formatted_message)
                send_button = self.driver.find_element(By.CSS_SELECTOR, ".msg-form__send-button")
                send_button.click()
                time.sleep(random.uniform(3, 5))  # Rate limiting

                close_button = self.driver.find_element(By.CSS_SELECTOR, "button[aria-label*='Dismiss']")
                close_button.click()
                time.sleep(1)

                self.log(f"Sent message to {contact['name']}")
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
