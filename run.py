import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import threading
import AutoMailerPro_v5_1
import os

def run_campaign():
    global selected_mode, sales_file_path, letter_content, subject_line
    selected_mode = mode_var.get()
    sales_file_path = file_entry.get()
    letter_content = letter_text.get("1.0", tk.END).strip() if not use_default_var.get() else None
    subject_line = subject_entry.get().strip()  # Get the subject line
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
        AutoMailerPro_v5_1.main(selected_mode, sales_file_path, letter_content, subject_line)
        root.after(0, update_ui_success)
    except Exception as err:
        root.after(0, lambda e=err: update_ui_error(str(e)))

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        file_entry.delete(0, tk.END)
        file_entry.insert(0, file_path)

def redirect_print(*args, **kwargs):
    output_text.insert(tk.END, ' '.join(map(str, args)) + '\n')
    output_text.see(tk.END)
    root.update()

def update_subject_line(*args):
    """Update the subject line entry based on the selected mode."""
    mode = mode_var.get()
    if not user_edited_subject.get():  # Only update if user hasn't edited the subject
        subject_entry.delete(0, tk.END)
        if mode == "personal":
            subject_entry.insert(0, "Homeowners Insurance Rates Are Finally on the Decline – Don’t Miss Out!")
        else:  # commercial
            subject_entry.insert(0, "Protect Your Business with Tailored Insurance Solutions!")

def mark_subject_edited(event):
    """Mark the subject line as user-edited when modified."""
    user_edited_subject.set(True)

import sys
sys.stdout.write = redirect_print
sys.stderr.write = redirect_print

root = tk.Tk()
root.title("Auto Mailer Pro    © 2025 Kyle Padilla — Jones Insurance Advisors, Inc.")
root.geometry("1000x825")

main_frame = ttk.Frame(root, padding="10")
main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

# Add logo 
logo_path = os.path.join("logo.png")  # Adjust path 
if os.path.exists(logo_path):
    logo_image = tk.PhotoImage(file=logo_path)
    # Scale down the image
    logo_image = logo_image.subsample(2, 2)  # Reduce size by factor of 2
    logo_label = ttk.Label(main_frame, image=logo_image)
    logo_label.image = logo_image  # Keep a reference to avoid garbage collection
    logo_label.grid(row=0, column=0, columnspan=4, pady=10)
else:
    print(f"❌ Logo file not found: {logo_path}")
    logo_label = ttk.Label(main_frame, text="Logo Not Found", font=("Arial", 12))
    logo_label.grid(row=0, column=0, columnspan=4, pady=10)

ttk.Label(main_frame, text="Select Mode:").grid(row=1, column=0, sticky=tk.W, pady=5)
mode_var = tk.StringVar(value="personal")
user_edited_subject = tk.BooleanVar(value=False)  # Track if subject line was edited
mode_var.trace("w", update_subject_line)  # Update subject line when mode changes
ttk.Radiobutton(main_frame, text="Personal Lines", variable=mode_var, value="personal").grid(row=1, column=1, sticky=tk.W)
ttk.Radiobutton(main_frame, text="Commercial Lines", variable=mode_var, value="commercial").grid(row=1, column=2, sticky=tk.W)

ttk.Label(main_frame, text="Sales Data File:").grid(row=2, column=0, sticky=tk.W, pady=5)
file_entry = ttk.Entry(main_frame, width=50)
file_entry.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
ttk.Button(main_frame, text="Browse", command=browse_file).grid(row=2, column=3, padx=5)

ttk.Label(main_frame, text="Subject Line:").grid(row=3, column=0, sticky=tk.W, pady=5)
subject_entry = ttk.Entry(main_frame, width=50)
subject_entry.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
subject_entry.insert(0, "Homeowners Insurance Rates Are Finally on the Decline – Don’t Miss Out!")  # Default for personal
subject_entry.bind("<KeyRelease>", mark_subject_edited)  # Mark as edited on key press

use_default_var = tk.BooleanVar(value=True)
ttk.Checkbutton(main_frame, text="Use Default Template", variable=use_default_var).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=5)

ttk.Label(main_frame, text="Letter Content (leave blank if using default):").grid(row=5, column=0, sticky=tk.W, pady=5)
letter_text = scrolledtext.ScrolledText(main_frame, width=60, height=10)
letter_text.grid(row=6, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)

run_button = ttk.Button(main_frame, text="Run Campaign", command=run_campaign)
run_button.grid(row=7, column=0, columnspan=4, pady=10)

progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
progress_bar.grid(row=8, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)

ttk.Label(main_frame, text="Output:").grid(row=9, column=0, sticky=tk.W, pady=5)
output_text = scrolledtext.ScrolledText(main_frame, width=60, height=10)
output_text.grid(row=10, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)

# Add credits at the bottom
credits_label = ttk.Label(
    main_frame,
    text="Support: scooby_rizz@protonmail.com   |  Repository: github.io/scoobyrizz-py   |  Last updated: 08/13/2025",
    font=("Arial", 10)
)
credits_label.grid(row=11, column=0, columnspan=4, pady=5)

root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
main_frame.columnconfigure(1, weight=1)

root.mainloop()
