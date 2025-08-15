from ttkthemes import ThemedTk
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import AutoMailerPro_v5_1
import os
import sys

class StdoutRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, message):
        self.text_widget.insert(tk.END, message)
        self.text_widget.see(tk.END)
        self.text_widget.update()

    def flush(self):
        pass

def run_campaign():
    global selected_mode, sales_file_path, letter_content, subject_line, signature_name, signature_title, signature_image, signature_email
    selected_mode = mode_var.get()
    sales_file_path = file_entry.get()
    letter_content = letter_text.get("1.0", tk.END).strip() if not use_default_var.get() else None
    subject_line = subject_entry.get().strip()
    signature_name, signature_title, signature_image, signature_email = signature_profiles[signature_var.get()]
    if not sales_file_path:
        messagebox.showerror("Error", "Please select a sales data file!")
        return
    if not use_default_var.get() and not letter_content:
        messagebox.showerror("Error", "Please enter letter content or select default template!")
        return
    if not subject_line:
        messagebox.showerror("Error", "Please enter a subject line!")
        return
    run_button.config(state='disabled')
    progress_bar.start()
    output_text.delete("1.0", tk.END)
    threading.Thread(target=threaded_main, daemon=True).start()

def threaded_main():
    def update_ui_success():
        messagebox.showinfo("Success", "Email campaign completed successfully!")
        progress_bar.stop()
        run_button.config(state='normal')

    def update_ui_error(error_msg):
        messagebox.showerror("Error", f"Failed to run campaign: {error_msg}")
        progress_bar.stop()
        run_button.config(state='normal')

    try:
        AutoMailerPro_v5_1.main(
            selected_mode, sales_file_path, letter_content, subject_line,
            signature_name=signature_name, signature_title=signature_title,
            signature_image=signature_image, signature_email=signature_email
        )
        root.after(0, update_ui_success)
    except Exception as err:
        root.after(0, lambda e=err: update_ui_error(str(e)))

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)

def update_subject_line(*args):
    mode = mode_var.get()
    if not user_edited_subject.get():
        subject_entry.delete(0, tk.END)
        if mode == "personal":
            subject_entry.insert(0, "Homeowners Insurance Rates Are Finally on the Decline – Don’t Miss Out!")
        else:
            subject_entry.insert(0, "Protect Your Business with Tailored Insurance Solutions!")

def mark_subject_edited(event):
    user_edited_subject.set(True)

# Initialize window with theme
root = ThemedTk(theme="arc")
root.title("Auto Mailer Pro    © 2025 Kyle Padilla — Jones Insurance Advisors, Inc.")
root.geometry("1000x825")
root.configure(bg="#f0f4f8")

main_frame = ttk.Frame(root, padding="20")
main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
main_frame.configure(style="Main.TFrame")

# Define styles
style = ttk.Style()
style.configure("Main.TFrame", background="#f0f4f8")
style.configure("TButton", font=("Arial", 12), padding=10)

# Add logo
logo_path = os.path.join("logo.png")
if os.path.exists(logo_path):
    logo_image = tk.PhotoImage(file=logo_path)
    logo_image = logo_image.subsample(2, 2)
    logo_label = tk.Label(main_frame, image=logo_image, bg="#f0f4f8")
    logo_label.image = logo_image
    logo_label.grid(row=0, column=0, columnspan=4, pady=20)
else:
    print(f"❌ Logo file not found: {logo_path}")
    logo_label = tk.Label(main_frame, text="Logo Not Found", font=("Arial", 12), bg="#f0f4f8")
    logo_label.grid(row=0, column=0, columnspan=4, pady=20)

# Define signature profiles (name, title, image, email)
signature_profiles = {
    "Brian Jones": ("Brian Jones", "Vice President", "signature_brian.png", "Brian@jonesia.com"),
    "Robert Jones": ("Robert Jones", "President", "signature_bob.png", "bob@jonesia.com"),
    "Kyle Padilla": ("Jane Doe", "Insurance Agent", "signature_kyle.png", "kyle@jonesia.com"),
    "Julie Siano": ("Julie Siano", "Julie Siano", "signature_jane.png", "Jane@jonesia.com"),
}

# Signature selection
signature_label = tk.Label(main_frame, text="Signature:", font=("Arial", 12), bg="#f0f4f8")
signature_label.grid(row=1, column=0, sticky=tk.W, pady=5)
signature_var = tk.StringVar(value="Brian Jones")
signature_dropdown = ttk.Combobox(main_frame, textvariable=signature_var, values=list(signature_profiles.keys()), state="readonly")
signature_dropdown.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)

# Mode selection
mode_label = tk.Label(main_frame, text="Select Mode:", font=("Arial", 12), bg="#f0f4f8")
mode_label.grid(row=2, column=0, sticky=tk.W, pady=5)
mode_var = tk.StringVar(value="personal")
user_edited_subject = tk.BooleanVar(value=False)
mode_var.trace("w", update_subject_line)
ttk.Radiobutton(main_frame, text="Personal Lines", variable=mode_var, value="personal").grid(row=2, column=1, sticky=tk.W)
ttk.Radiobutton(main_frame, text="Commercial Lines", variable=mode_var, value="commercial").grid(row=2, column=2, sticky=tk.W)

# File selection
file_label = tk.Label(main_frame, text="Sales Data File:", font=("Arial", 12), bg="#f0f4f8")
file_label.grid(row=3, column=0, sticky=tk.W, pady=5)
file_entry = ttk.Entry(main_frame, width=50, font=("Arial", 10))
file_entry.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
ttk.Button(main_frame, text="Browse", command=browse_file, style="TButton").grid(row=3, column=3, padx=5)

# Subject line
subject_label = tk.Label(main_frame, text="Subject Line:", font=("Arial", 12), bg="#f0f4f8")
subject_label.grid(row=4, column=0, sticky=tk.W, pady=5)
subject_entry = ttk.Entry(main_frame, width=50, font=("Arial", 10))
subject_entry.grid(row=4, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
subject_entry.insert(0, "Homeowners Insurance Rates Are Finally on the Decline – Don’t Miss Out!")
subject_entry.bind("<KeyRelease>", mark_subject_edited)

# Default template checkbox
use_default_var = tk.BooleanVar(value=True)
default_check = ttk.Checkbutton(main_frame, text="Use Default Template", variable=use_default_var)
default_check.grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=5)

# Letter content
content_label = tk.Label(main_frame, text="Letter Content (leave blank if using default):", font=("Arial", 12), bg="#f0f4f8")
content_label.grid(row=6, column=0, sticky=tk.W, pady=5)
letter_text = scrolledtext.ScrolledText(main_frame, width=60, height=10, font=("Arial", 10), bg="white", fg="black")
letter_text.grid(row=7, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)

# Run button
run_button = ttk.Button(main_frame, text="Run Campaign", command=run_campaign, style="TButton")
run_button.grid(row=8, column=0, columnspan=4, pady=20)

# Progress bar
progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
progress_bar.grid(row=9, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)

# Output text
output_label = tk.Label(main_frame, text="Output:", font=("Arial", 12), bg="#f0f4f8")
output_label.grid(row=10, column=0, sticky=tk.W, pady=5)
output_text = scrolledtext.ScrolledText(main_frame, width=60, height=10, font=("Arial", 10), bg="white", fg="black")
output_text.grid(row=11, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)

# Redirect print output to GUI
sys.stdout = StdoutRedirector(output_text)
sys.stderr = StdoutRedirector(output_text)

# Credits
credits_label = tk.Label(
    main_frame,
    text="Support: scooby_rizz@protonmail.com   |  Repository: github.io/scoobyrizz-py   |  Last updated: 08/14/2025",
    font=("Arial", 10),
    bg="#f0f4f8"
)
credits_label.grid(row=12, column=0, columnspan=4, pady=20)

root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
main_frame.columnconfigure(1, weight=1)

root.mainloop()