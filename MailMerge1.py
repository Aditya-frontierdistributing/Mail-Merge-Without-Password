import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import smtplib
from email.message import EmailMessage
import mimetypes
import os

class MailMergeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Mail Merge Application")

        self.csv_file_path = ""

        # Create GUI components
        self.create_widgets()

    def create_widgets(self):
        # Email address
        tk.Label(self.root, text="Email Address:").grid(row=0, column=0, padx=10, pady=5)
        self.email_entry = tk.Entry(self.root, width=40)
        self.email_entry.grid(row=0, column=1, padx=10, pady=5)

        # Select CSV file
        tk.Label(self.root, text="CSV File:").grid(row=1, column=0, padx=10, pady=5)
        self.csv_label = tk.Label(self.root, text="No file selected")
        self.csv_label.grid(row=1, column=1, padx=10, pady=5)
        self.select_csv_button = tk.Button(self.root, text="Browse", command=self.select_csv_file)
        self.select_csv_button.grid(row=1, column=2, padx=10, pady=5)

        # Email subject
        tk.Label(self.root, text="Email Subject:").grid(row=2, column=0, padx=10, pady=5)
        self.subject_entry = tk.Entry(self.root, width=40)
        self.subject_entry.grid(row=2, column=1, padx=10, pady=5)

        # Email body
        tk.Label(self.root, text="Email Body:").grid(row=3, column=0, padx=10, pady=5)
        self.body_text = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, width=50, height=10)
        self.body_text.insert(tk.END, "Dear {first_name} {last_name},\n\n")
        self.body_text.grid(row=3, column=1, padx=10, pady=5, columnspan=2)

        # Send Email button
        self.send_button = tk.Button(self.root, text="Send Emails", command=self.send_emails)
        self.send_button.grid(row=4, column=1, pady=20)

    def select_csv_file(self):
        self.csv_file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        self.csv_label.config(text=os.path.basename(self.csv_file_path))

    def send_emails(self):
        if not self.csv_file_path:
            messagebox.showerror("Error", "Please select a CSV file.")
            return

        if not self.email_entry.get():
            messagebox.showerror("Error", "Please enter an email address.")
            return

        if not self.subject_entry.get():
            messagebox.showerror("Error", "Please enter an email subject.")
            return

        # Read the CSV file
        df = pd.read_csv(self.csv_file_path)

        email_address = self.email_entry.get()
        email_subject = self.subject_entry.get()
        email_body_template = self.body_text.get("1.0", tk.END)

        # Identify attachment columns
        attachment_columns = [col for col in df.columns if col.startswith('Attachment')]

        # Email sending function
        def send_mail(to_email, first_name, last_name, cc, bcc, attachments):
            msg = EmailMessage()
            msg['From'] = email_address
            msg['To'] = to_email
            msg['Subject'] = email_subject

            # Add Cc and Bcc fields if they are not NaN
            if pd.notna(cc):
                msg['Cc'] = cc
            if pd.notna(bcc):
                msg['Bcc'] = bcc

            body = email_body_template.format(first_name=first_name, last_name=last_name)
            msg.set_content(body)

            # Attach files
            for attachment in attachments:
                if pd.notna(attachment) and os.path.isfile(attachment):
                    mime_type, _ = mimetypes.guess_type(attachment)
                    if mime_type:
                        mime_type, mime_subtype = mime_type.split('/')
                        with open(attachment, 'rb') as file:
                            msg.add_attachment(file.read(),
                                               maintype=mime_type,
                                               subtype=mime_subtype,
                                               filename=os.path.basename(attachment))

            try:
                with smtplib.SMTP('10.0.0.170', 25, timeout=30) as smtp:
                    smtp.ehlo()  # Can be omitted
                    smtp.send_message(msg)
                    print(f"Email sent to {to_email}")
                    return True
            except Exception as e:
                print(f"Failed to send email to {to_email}: {e}")
                return False

        # Iterate through the dataframe and send emails
        all_sent = True
        for index, row in df.iterrows():
            attachments = [row[col] for col in attachment_columns]
            if not send_mail(row['email'], row['first_name'], row['last_name'], row.get('cc'), row.get('bcc'), attachments):
                all_sent = False

        if all_sent:
            messagebox.showinfo("Success", "Emails sent successfully!")
        else:
            messagebox.showwarning("Warning", "Some emails could not be sent. Check the logs for details.")

if __name__ == "__main__":
    root = tk.Tk()
    app = MailMergeApp(root)
    root.mainloop()