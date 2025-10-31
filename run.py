from pathlib import Path
import sys
import threading


import AutoMailerPro
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from ttkthemes import ThemedTk
from textwrap import dedent

def get_base_dir() -> Path:
    """Return the directory that holds bundled resources."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent


BASE_DIR = get_base_dir()
ASSETS_DIR = BASE_DIR / "assets"
SIGNATURES_DIR = ASSETS_DIR / "signatures"

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
    selected_template = template_var.get()
    if selected_template == "Custom":
        letter_content = letter_text.get("1.0", tk.END).strip()
        if not letter_content:
            messagebox.showerror("Error", "Please enter letter content for the custom template!")
            return
    else:
        letter_content = LETTER_TEMPLATES[selected_template][selected_mode]
    subject_line = subject_entry.get().strip()
    signature_name, signature_title, signature_image, signature_email = signature_profiles[signature_var.get()]
    if not sales_file_path:
        messagebox.showerror("Error", "Please select a sales data file!")
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
        AutoMailerPro.main(
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
    if template_var.get() != "Custom":
        apply_template_selection()

def mark_subject_edited(event):
    user_edited_subject.set(True)

def apply_template_selection(*args):
    global custom_content_cache, current_template_selection
    new_selection = template_var.get()
    mode = mode_var.get()

    if current_template_selection == "Custom" and new_selection != "Custom":
        custom_content_cache = letter_text.get("1.0", tk.END).strip()

    if new_selection == "Custom":
        letter_text.config(state='normal')
        if current_template_selection != "Custom":
            letter_text.delete("1.0", tk.END)
            if custom_content_cache:
                letter_text.insert(tk.END, custom_content_cache)
        current_template_selection = new_selection
        return

    template_content = LETTER_TEMPLATES[new_selection][mode]
    letter_text.config(state='normal')
    letter_text.delete("1.0", tk.END)
    letter_text.insert(tk.END, template_content)
    letter_text.config(state='disabled')
    current_template_selection = new_selection

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

logo_path = ASSETS_DIR / "logo.png"
if logo_path.exists():
    logo_image = tk.PhotoImage(file=str(logo_path))
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
"Brian Jones": (
        "Brian Jones",
        "Vice President",
        SIGNATURES_DIR / "signature_brian.png",
        "Brian@jonesia.com",
    ),
    "Robert Jones": (
        "Robert Jones",
        "President",
        SIGNATURES_DIR / "signature_bob.png",
        "bob@jonesia.com",
    ),
    "Kyle Padilla": (
        "Kyle Padilla",
        "Insurance Agent",
        SIGNATURES_DIR / "signature_kyle.png",
        "kyle@jonesia.com",
    ),
    "Kristofer Siggins": (
        "Kristofer Siggins",
        "Account Executive",
        SIGNATURES_DIR / "signature_kris.png",
        "Kris@jonesia.com",
    ),
}

INDIAN_RIVER_PERSONAL_TEMPLATE = dedent(
    """
For the first time in years, homeowners rates are coming down — and the savings could be significant.

Recent legislative changes have boosted competition in Florida’s property insurance market, and many Indian River County homeowners are already benefiting.

Jones Insurance Advisors is a two-generation, family-owned independent agency located right here in Vero Beach. Our team of dedicated agents possess extensive knowledge of the intricacies of the local insurance market, and are excited to assist you in finding the most comprehensive and competitively priced insurance solutions.

Call us today for a free, no-obligation quote, or visit our website below and complete a quote request, and one of our dedicated agents will reach out to you!

We look forward to earning your business and providing you the personal, dedicated service you have come to expect by doing business locally.

Warm Regards,
"""
).strip()

INDIAN_RIVER_COMMERCIAL_TEMPLATE = dedent(
    """
Protecting your business is our priority at Jones Insurance Advisors.

As an Indian River County business, you need insurance solutions tailored to your unique needs. Our experienced team specializes in crafting comprehensive coverage plans for businesses like yours, ensuring protection against risks while keeping costs competitive.

Jones Insurance Advisors, a family-owned agency in Vero Beach, is here to help. Contact us for a free consultation to discuss how we can safeguard your business.

We look forward to partnering with you!

Best Regards,
"""
).strip()

ST_LUCIE_PERSONAL_TEMPLATE = dedent(
    """
For the first time in years, homeowners rates are coming down — and the savings could be significant.

Recent legislative changes have boosted competition in Florida’s property insurance market, and many St. Lucie County homeowners are already benefiting.

Jones Insurance Advisors is a two-generation, family-owned independent agency located right here on the Treasure Coast. Our team of dedicated agents possess extensive knowledge of the intricacies of the local insurance market, and are excited to assist you in finding the most comprehensive and competitively priced insurance solutions.

Call us today for a free, no-obligation quote, or visit our website below and complete a quote request, and one of our dedicated agents will reach out to you!

We look forward to earning your business and providing you the personal, dedicated service you have come to expect by doing business locally.

Warm Regards,
"""
).strip()

ST_LUCIE_COMMERCIAL_TEMPLATE = dedent(
    """
Jones Insurance Advisors is focused on protecting St. Lucie County businesses like yours.

Whether you’re operating in Port St. Lucie, Fort Pierce, or along the coast, you need insurance solutions built around the unique exposures your company faces.

Our experienced advisors craft comprehensive coverage portfolios that balance protection and cost, so you can stay focused on growing your business.

As a family-owned independent agency serving the entire Treasure Coast, we’re ready to connect and explore how we can safeguard your operations.

We look forward to partnering with you!

Best Regards,
"""
).strip()

LETTER_TEMPLATES = {
    "Indian River County": {
        "personal": INDIAN_RIVER_PERSONAL_TEMPLATE,
        "commercial": INDIAN_RIVER_COMMERCIAL_TEMPLATE,
    },
    "St. Lucie County": {
        "personal": ST_LUCIE_PERSONAL_TEMPLATE,
        "commercial": ST_LUCIE_COMMERCIAL_TEMPLATE,
    },
}

custom_content_cache = ""
current_template_selection = None
customer_window = None


def open_customer_manager():
    """Display a window for managing customers and viewing reports."""

    global customer_window
    if customer_window and tk.Toplevel.winfo_exists(customer_window):
        customer_window.deiconify()
        customer_window.lift()
        customer_window.focus_force()
        return

    customer_window = tk.Toplevel(root)
    customer_window.title("Customer Database & Reports")
    customer_window.geometry("950x620")
    customer_window.configure(bg="#f0f4f8")

    container = ttk.Frame(customer_window, padding="20")
    container.pack(fill=tk.BOTH, expand=True)
    container.columnconfigure(0, weight=1)
    container.rowconfigure(0, weight=1)

    columns = (
        "name",
        "email",
        "phone",
        "premium",
        "home_price",
        "responded",
        "converted",
    )
    headings = {
        "name": "Name",
        "email": "Email",
        "phone": "Phone",
        "premium": "Premium",
        "home_price": "Home Price",
        "responded": "Responded",
        "converted": "Converted",
    }

    tree = ttk.Treeview(container, columns=columns, show="headings", height=10)
    tree.grid(row=0, column=0, columnspan=4, sticky="nsew")

    for column in columns:
        tree.heading(column, text=headings[column])
        anchor = tk.W if column not in {"premium", "home_price"} else tk.E
        width = 160 if column == "name" else 130
        tree.column(column, width=width, anchor=anchor)

    scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=0, column=4, sticky="ns")

    form_frame = ttk.LabelFrame(container, text="Customer Details", padding="15")
    form_frame.grid(row=1, column=0, columnspan=5, sticky="ew", pady=(15, 0))
    for index in range(4):
        form_frame.columnconfigure(index, weight=1)

    name_var = tk.StringVar()
    email_var = tk.StringVar()
    phone_var = tk.StringVar()
    premium_var = tk.StringVar()
    home_price_var = tk.StringVar()
    responded_var = tk.BooleanVar()
    converted_var = tk.BooleanVar()
    selected_customer = {"id": None}
    customer_rows = []

    ttk.Label(form_frame, text="Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
    ttk.Entry(form_frame, textvariable=name_var).grid(row=0, column=1, sticky="ew", pady=5)

    ttk.Label(form_frame, text="Email:").grid(row=0, column=2, sticky=tk.W, pady=5)
    ttk.Entry(form_frame, textvariable=email_var).grid(row=0, column=3, sticky="ew", pady=5)

    ttk.Label(form_frame, text="Phone:").grid(row=1, column=0, sticky=tk.W, pady=5)
    ttk.Entry(form_frame, textvariable=phone_var).grid(row=1, column=1, sticky="ew", pady=5)

    ttk.Label(form_frame, text="Premium:").grid(row=1, column=2, sticky=tk.W, pady=5)
    ttk.Entry(form_frame, textvariable=premium_var).grid(row=1, column=3, sticky="ew", pady=5)

    ttk.Label(form_frame, text="Home Price:").grid(row=2, column=0, sticky=tk.W, pady=5)
    ttk.Entry(form_frame, textvariable=home_price_var).grid(row=2, column=1, sticky="ew", pady=5)

    ttk.Checkbutton(
        form_frame,
        text="Customer Responded",
        variable=responded_var,
    ).grid(row=2, column=2, sticky=tk.W, pady=5)

    ttk.Checkbutton(
        form_frame,
        text="Converted to Policyholder",
        variable=converted_var,
    ).grid(row=2, column=3, sticky=tk.W, pady=5)

    def clear_form():
        name_var.set("")
        email_var.set("")
        phone_var.set("")
        premium_var.set("")
        home_price_var.set("")
        responded_var.set(False)
        converted_var.set(False)
        selected_customer["id"] = None
        tree.selection_remove(tree.selection())

    def refresh_tree():
        nonlocal customer_rows
        for item in tree.get_children():
            tree.delete(item)
        try:
            customer_rows = AutoMailerPro.list_customers()
        except Exception as exc:
            messagebox.showerror("Error", f"Unable to load customers: {exc}")
            customer_rows = []
            return

        for customer in customer_rows:
            try:
                premium_value = float(customer.get("premium") or 0)
            except (TypeError, ValueError):
                premium_value = 0.0
            try:
                home_price_value = float(customer.get("home_price") or 0)
            except (TypeError, ValueError):
                home_price_value = 0.0

            tree.insert(
                "",
                tk.END,
                iid=str(customer["id"]),
                values=(
                    customer.get("name", ""),
                    customer.get("email", ""),
                    customer.get("phone", ""),
                    f"${premium_value:,.2f}",
                    f"${home_price_value:,.0f}",
                    "Yes" if customer.get("responded") else "No",
                    "Yes" if customer.get("converted") else "No",
                ),
            )

    def on_select(event):
        selection = tree.selection()
        if not selection:
            return
        customer_id = selection[0]
        customer = next(
            (row for row in customer_rows if str(row.get("id")) == customer_id),
            None,
        )
        if not customer:
            return

        selected_customer["id"] = customer.get("id")
        name_var.set(customer.get("name", ""))
        email_var.set(customer.get("email", ""))
        phone_var.set(customer.get("phone", ""))
        premium_var.set(str(customer.get("premium", "") or ""))
        home_price_var.set(str(customer.get("home_price", "") or ""))
        responded_var.set(bool(customer.get("responded")))
        converted_var.set(bool(customer.get("converted")))

    def save_selected_customer():
        payload = {
            "id": selected_customer.get("id"),
            "name": name_var.get().strip(),
            "email": email_var.get().strip(),
            "phone": phone_var.get().strip(),
            "premium": premium_var.get().strip(),
            "home_price": home_price_var.get().strip(),
            "responded": responded_var.get(),
            "converted": converted_var.get(),
        }
        try:
            saved_id = AutoMailerPro.save_customer(payload)
        except ValueError as exc:
            messagebox.showerror("Validation Error", str(exc))
            return
        except Exception as exc:
            messagebox.showerror("Error", f"Unable to save customer: {exc}")
            return

        action = "updated" if payload.get("id") else "added"
        messagebox.showinfo("Success", f"Customer {action} successfully.")
        selected_customer["id"] = saved_id
        refresh_tree()
        clear_form()

    def show_report():
        try:
            metrics = AutoMailerPro.get_customer_metrics()
        except Exception as exc:
            messagebox.showerror("Error", f"Unable to generate report: {exc}")
            return

        report_popup = tk.Toplevel(customer_window)
        report_popup.title("Customer Engagement Report")
        report_popup.geometry("400x260")
        report_popup.configure(bg="#f0f4f8")
        report_popup.transient(customer_window)
        report_popup.grab_set()
        report_popup.focus_set()

        ttk.Label(
            report_popup,
            text="Customer Engagement Summary",
            font=("Arial", 14, "bold"),
            padding=10,
        ).pack()

        content = ttk.Frame(report_popup, padding="10")
        content.pack(fill=tk.BOTH, expand=True)
        stats = [
            ("Total Customers", metrics.get("total_customers", 0)),
            (
                "Responses",
                f"{metrics.get('responded_count', 0)} ({metrics.get('response_rate', 0.0):.1f}%)",
            ),
            (
                "Conversions",
                f"{metrics.get('converted_count', 0)} ({metrics.get('conversion_rate', 0.0):.1f}%)",
            ),
            (
                "Average Premium",
                f"${metrics.get('average_premium', 0.0):,.2f}",
            ),
            (
                "Average Home Price",
                f"${metrics.get('average_home_price', 0.0):,.0f}",
            ),
        ]

        for idx, (label, value) in enumerate(stats):
            ttk.Label(content, text=label + ":", font=("Arial", 11, "bold")).grid(
                row=idx, column=0, sticky=tk.W, pady=5
            )
            ttk.Label(content, text=value, font=("Arial", 11)).grid(
                row=idx, column=1, sticky=tk.W, pady=5
            )

    button_frame = ttk.Frame(container, padding="5")
    button_frame.grid(row=2, column=0, columnspan=5, sticky="ew", pady=(15, 0))
    button_frame.columnconfigure(0, weight=1)
    button_frame.columnconfigure(1, weight=1)
    button_frame.columnconfigure(2, weight=1)

    ttk.Button(button_frame, text="Save Customer", command=save_selected_customer).grid(
        row=0, column=0, sticky="ew", padx=5
    )
    ttk.Button(button_frame, text="Clear", command=clear_form).grid(
        row=0, column=1, sticky="ew", padx=5
    )
    ttk.Button(button_frame, text="Show Report", command=show_report).grid(
        row=0, column=2, sticky="ew", padx=5
    )

    tree.bind("<<TreeviewSelect>>", on_select)

    def on_close():
        nonlocal customer_rows
        global customer_window
        customer_rows = []
        clear_form()
        if customer_window is not None:
            window_to_close = customer_window
            customer_window = None
            window_to_close.destroy()

    customer_window.protocol("WM_DELETE_WINDOW", on_close)

    refresh_tree()


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
template_var = tk.StringVar(value="Indian River County")
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

# Template selection
template_label = tk.Label(main_frame, text="Letter Template:", font=("Arial", 12), bg="#f0f4f8")
template_label.grid(row=4, column=0, sticky=tk.W, pady=5)
template_dropdown = ttk.Combobox(
    main_frame,
    textvariable=template_var,
    values=list(LETTER_TEMPLATES.keys()) + ["Custom"],
    state="readonly"
)
template_dropdown.grid(row=4, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)

# Subject line
subject_label = tk.Label(main_frame, text="Subject Line:", font=("Arial", 12), bg="#f0f4f8")
subject_label.grid(row=5, column=0, sticky=tk.W, pady=5)
subject_entry = ttk.Entry(main_frame, width=50, font=("Arial", 10))
subject_entry.grid(row=5, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
subject_entry.insert(0, "Homeowners Insurance Rates Are Finally on the Decline – Don’t Miss Out!")
subject_entry.bind("<KeyRelease>", mark_subject_edited)

# Letter content
content_label = tk.Label(main_frame, text="Letter Content (editable when using the custom template):", font=("Arial", 12), bg="#f0f4f8")
content_label.grid(row=6, column=0, sticky=tk.W, pady=5)
letter_text = scrolledtext.ScrolledText(main_frame, width=60, height=10, font=("Arial", 10), bg="white", fg="black")
letter_text.grid(row=7, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)
template_var.trace_add("write", lambda *args: apply_template_selection())
apply_template_selection()
# Customer manager button
customer_manager_button = ttk.Button(
    main_frame,
    text="Customer Database & Reports",
    command=open_customer_manager,
    style="TButton",
)
customer_manager_button.grid(row=8, column=0, columnspan=4, pady=(15, 5))

# Run button
run_button = ttk.Button(main_frame, text="Run Campaign", command=run_campaign, style="TButton")
run_button.grid(row=9, column=0, columnspan=4, pady=20)

# Progress bar
progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
progress_bar.grid(row=10, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)

# Output text
output_label = tk.Label(main_frame, text="Output:", font=("Arial", 12), bg="#f0f4f8")
output_label.grid(row=11, column=0, sticky=tk.W, pady=5)
output_text = scrolledtext.ScrolledText(main_frame, width=60, height=10, font=("Arial", 10), bg="white", fg="black")
output_text.grid(row=12, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)

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
credits_label.grid(row=13, column=0, columnspan=4, pady=20)

root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
main_frame.columnconfigure(1, weight=1)

if __name__ == "__main__":
    root.mainloop()