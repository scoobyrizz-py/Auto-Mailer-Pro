from pathlib import Path
import json
import re
import shutil
import sys
import threading

from fuzzywuzzy import fuzz

import AutoMailerPro
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
from ttkthemes import ThemedTk
from textwrap import dedent
from typing import Dict, List, Optional

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
DEFAULT_SIGNATURE_PROFILES = {
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
CUSTOM_SIGNATURES_DIR = AutoMailerPro.WRITABLE_DATA_DIR / "signatures"
CUSTOM_SIGNATURES_FILE = AutoMailerPro.WRITABLE_DATA_DIR / "signature_profiles.json"

signature_profiles = dict(DEFAULT_SIGNATURE_PROFILES)


def sanitize_filename(label: str) -> str:
    """Return a filesystem-safe filename derived from ``label``."""

    sanitized = re.sub(r"[^A-Za-z0-9_-]", "_", label.strip())
    return sanitized or "signature"


def load_custom_signatures() -> Dict[str, tuple]:
    """Load persisted signature profiles from the user data directory."""

    CUSTOM_SIGNATURES_DIR.mkdir(parents=True, exist_ok=True)
    if not CUSTOM_SIGNATURES_FILE.exists():
        return {}

    try:
        with CUSTOM_SIGNATURES_FILE.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
    except (OSError, json.JSONDecodeError) as exc:
        print(f"⚠️ Unable to load custom signatures: {exc}")
        return {}

    loaded_profiles: Dict[str, tuple] = {}
    for entry in payload:
        name = entry.get("name")
        title = entry.get("title")
        email = entry.get("email")
        image_name = entry.get("image")
        if not name:
            continue
        image_path = None
        if image_name:
            candidate = Path(image_name)
            image_path = candidate if candidate.is_absolute() else CUSTOM_SIGNATURES_DIR / candidate
        loaded_profiles[name] = (
            name,
            title or "",
            image_path,
            email or "",
        )

    return loaded_profiles


def persist_custom_signatures() -> None:
    """Write custom signature definitions to disk."""

    entries = []
    for name, profile in signature_profiles.items():
        if name in DEFAULT_SIGNATURE_PROFILES:
            continue
        _, title, image_path, email = profile
        relative_image = None
        if image_path:
            try:
                image_path = Path(image_path)
                if image_path.is_relative_to(CUSTOM_SIGNATURES_DIR):
                    relative_image = image_path.relative_to(CUSTOM_SIGNATURES_DIR).as_posix()
                else:
                    relative_image = image_path.as_posix()
            except AttributeError:
                if str(image_path).startswith(str(CUSTOM_SIGNATURES_DIR)):
                    relative_image = str(image_path)[len(str(CUSTOM_SIGNATURES_DIR)) + 1 :]
                else:
                    relative_image = str(image_path)

        entries.append(
            {
                "name": name,
                "title": title,
                "email": email,
                "image": relative_image,
            }
        )

    try:
        CUSTOM_SIGNATURES_FILE.parent.mkdir(parents=True, exist_ok=True)
        with CUSTOM_SIGNATURES_FILE.open("w", encoding="utf-8") as handle:
            json.dump(entries, handle, indent=2)
    except OSError as exc:
        messagebox.showerror("Error", f"Unable to save signature profiles: {exc}")


signature_profiles.update(load_custom_signatures())

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
    customer_window.geometry("1020x680")
    customer_window.configure(bg="#f0f4f8")

    container = ttk.Frame(customer_window, padding="20")
    container.pack(fill=tk.BOTH, expand=True)
    container.columnconfigure(0, weight=1)
    container.rowconfigure(1, weight=1)

    search_frame = ttk.LabelFrame(container, text="Search & Quick Find", padding="15")
    search_frame.grid(row=0, column=0, columnspan=5, sticky="ew")
    search_frame.columnconfigure(0, weight=1)
    search_frame.columnconfigure(1, weight=0)
    search_frame.columnconfigure(2, weight=0)

    search_var = tk.StringVar()
    search_entry = ttk.Entry(search_frame, textvariable=search_var)
    search_entry.grid(row=0, column=0, sticky="ew", pady=5)
    ttk.Button(search_frame, text="Search", command=lambda: perform_search()).grid(
        row=0, column=1, padx=5, pady=5
    )
    ttk.Button(search_frame, text="Clear Search", command=lambda: clear_search()).grid(
        row=0, column=2, padx=5, pady=5
    )
    ttk.Label(
        search_frame,
        text="Type a first name to auto-complete matching contacts.",
        font=("Arial", 9),
    ).grid(row=1, column=0, columnspan=2, sticky=tk.W)
    suggestion_var = tk.StringVar(value=())
    suggestion_listbox = tk.Listbox(
        search_frame,
        listvariable=suggestion_var,
        height=5,
        exportselection=False,
    )
    suggestion_listbox.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(4, 0))
    suggestion_listbox.grid_remove()
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

    tree = ttk.Treeview(container, columns=columns, show="headings", height=12)
    tree.grid(row=1, column=0, columnspan=4, sticky="nsew")

    for column in columns:
        tree.heading(column, text=headings[column])
        anchor = tk.W if column not in {"premium", "home_price"} else tk.E
        width = 180 if column == "name" else 150
        tree.column(column, width=width, anchor=anchor)

    scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.grid(row=1, column=4, sticky="ns")

    report_options_frame = ttk.LabelFrame(container, text="Report Filters", padding="15")
    report_options_frame.grid(row=2, column=0, columnspan=5, sticky="ew", pady=(15, 0))

    include_prospects_var = tk.BooleanVar(value=True)
    include_responded_var = tk.BooleanVar(value=True)
    include_converted_var = tk.BooleanVar(value=True)

    ttk.Label(report_options_frame, text="Premium Tracking:").grid(
        row=0, column=0, sticky=tk.W, pady=5
    )
    ttk.Checkbutton(
        report_options_frame,
        text="Prospects",
        variable=include_prospects_var,
    ).grid(row=0, column=1, sticky=tk.W, padx=5, pady=5)
    ttk.Checkbutton(
        report_options_frame,
        text="Responded",
        variable=include_responded_var,
    ).grid(row=0, column=2, sticky=tk.W, padx=5, pady=5)
    ttk.Checkbutton(
        report_options_frame,
        text="Converted Clients",
        variable=include_converted_var,
    ).grid(row=0, column=3, sticky=tk.W, padx=5, pady=5)
    ttk.Button(
        report_options_frame,
        text="Generate Report",
        command=lambda: show_report(),
    ).grid(row=1, column=0, columnspan=4, sticky=tk.E, pady=(10, 0))

    form_frame = ttk.LabelFrame(container, text="Customer Details", padding="15")
    form_frame.grid(row=3, column=0, columnspan=5, sticky="ew", pady=(15, 0))
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
    customer_rows: list = []
    name_index: list = []
    current_search_matches: Optional[List[Dict[str, object]]] = None
    search_trace_id = None
    suggestions_suppressed = False
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

    def clear_suggestions():
        suggestion_var.set(())
        suggestion_listbox.selection_clear(0, tk.END)
        suggestion_listbox.grid_remove()

    def update_suggestion_box(prefix: str):
        nonlocal suggestions_suppressed
        if suggestions_suppressed:
            return
        if not prefix:
            clear_suggestions()
            return
        lowered = prefix.lower()
        matches = [
            candidate
            for candidate in name_index
            if candidate.lower().startswith(lowered)
        ][:8]
        if matches:
            suggestion_var.set(matches)
            suggestion_listbox.grid()
            suggestion_listbox.selection_clear(0, tk.END)
        else:
            clear_suggestions()


    def clear_search():
        nonlocal current_search_matches
        current_search_matches = None
        search_var.set("")
        search_entry.focus_set()
        clear_suggestions()

    def apply_filters():
        query = search_var.get().strip().lower()
        tree.delete(*tree.get_children())
        dataset = customer_rows if current_search_matches is None else current_search_matches
        for customer in dataset:
            searchable = " ".join(
                [
                    str(customer.get("name", "")),
                    str(customer.get("email", "")),
                    str(customer.get("phone", "")),
                ]
            ).lower()
            if current_search_matches is None and query and query not in searchable:
                continue          
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
                iid=str(customer.get("id")),
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
    def refresh_tree():
        nonlocal customer_rows, name_index, current_search_matches
        try:
            customer_rows = AutoMailerPro.list_customers()
        except Exception as exc:
            messagebox.showerror("Error", f"Unable to load customers: {exc}")
            customer_rows = []
            name_index = []
            current_search_matches = None
            tree.delete(*tree.get_children())
            return

        name_index = sorted(
            {
                str(row.get("name", "")).strip()
                for row in customer_rows
                if str(row.get("name", "")).strip()
            }
        )
        current_search_matches = None
        apply_filters()

        update_suggestion_box(search_var.get().strip())

    def focus_first_result(event=None):
        items = tree.get_children()
        if not items:
            return
        first = items[0]
        tree.selection_set(first)
        tree.focus(first)
        tree.see(first)
        on_select(None)

    def use_suggestion(event=None):
        nonlocal suggestions_suppressed
        selection = suggestion_listbox.curselection()
        if not selection:
            return "break"
        choice = suggestion_listbox.get(selection[0])
        suggestions_suppressed = True
        try:
            search_var.set(choice)
        finally:
            suggestions_suppressed = False
        suggestion_listbox.selection_clear(0, tk.END)
        suggestion_listbox.grid_remove()
        search_entry.focus_set()
        focus_first_result()
        return "break"

    def handle_search_key(event):
        if event.keysym == "Down" and suggestion_listbox.winfo_ismapped():
            suggestion_listbox.focus_set()
            if suggestion_listbox.size():
                suggestion_listbox.selection_set(0)
            return "break"
        if event.keysym == "Escape":
            clear_suggestions()
        return None

    def handle_suggestion_navigation(event):
        if event.keysym == "Escape":
            clear_suggestions()
            search_entry.focus_set()
            return "break"
        if event.keysym == "Up" and suggestion_listbox.curselection() == (0,):
            suggestion_listbox.selection_clear(0, tk.END)
            search_entry.focus_set()
            return "break"
        return None

    search_entry.bind("<KeyRelease>", handle_search_key)
    suggestion_listbox.bind("<Double-Button-1>", use_suggestion)
    suggestion_listbox.bind("<Return>", use_suggestion)
    suggestion_listbox.bind("<KP_Enter>", use_suggestion)
    suggestion_listbox.bind("<KeyRelease>", handle_suggestion_navigation)

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
        filters = {
            "include_prospects": include_prospects_var.get(),
            "include_responded": include_responded_var.get(),
            "include_converted": include_converted_var.get(),
        }
        if not any(filters.values()):
            messagebox.showwarning(
                "Report Filters",
                "Select at least one premium tracking option before generating a report.",
            )
            return        
        try:
            metrics = AutoMailerPro.get_customer_metrics(filters)
        except Exception as exc:
            messagebox.showerror("Error", f"Unable to generate report: {exc}")
            return

        report_popup = tk.Toplevel(customer_window)
        report_popup.title("Customer Engagement Report")
        report_popup.geometry("520x420")
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

        overall_stats = [
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
                        (
                "Total Premium",
                f"${metrics.get('total_premium', 0.0):,.2f}",
            ),
        ]

        overall_frame = ttk.LabelFrame(content, text="Overall Totals", padding="10")
        overall_frame.pack(fill=tk.X, pady=(0, 10))
        for idx, (label, value) in enumerate(overall_stats):
            ttk.Label(overall_frame, text=label + ":", font=("Arial", 11, "bold")).grid(
                row=idx, column=0, sticky=tk.W, pady=3
            )
            ttk.Label(overall_frame, text=value, font=("Arial", 11)).grid(
                row=idx, column=1, sticky=tk.W, pady=3
            )

        filtered = metrics.get("filtered_totals", {})
        filtered_frame = ttk.LabelFrame(content, text="Filtered Selection", padding="10")
        filtered_frame.pack(fill=tk.X, pady=(0, 10))

        active_filters = [
            label
            for label, enabled in (
                ("Prospects", filters["include_prospects"]),
                ("Responded", filters["include_responded"]),
                ("Converted", filters["include_converted"]),
            )
            if enabled
        ]
        filters_text = ", ".join(active_filters) if active_filters else "None"
        ttk.Label(
            filtered_frame,
            text=f"Filters Applied: {filters_text}",
            font=("Arial", 10, "italic"),
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=(0, 8))

        filtered_stats = [
            ("Customers Included", filtered.get("total_customers", 0)),
            ("Response Rate", f"{filtered.get('response_rate', 0.0):.1f}%"),
            ("Conversion Rate", f"{filtered.get('conversion_rate', 0.0):.1f}%"),
            ("Total Premium", f"${filtered.get('total_premium', 0.0):,.2f}"),
            ("Average Premium", f"${filtered.get('average_premium', 0.0):,.2f}"),
        ]

        for idx, (label, value) in enumerate(filtered_stats, start=1):
            ttk.Label(filtered_frame, text=label + ":", font=("Arial", 11, "bold")).grid(
                row=idx, column=0, sticky=tk.W, pady=3
            )
            ttk.Label(filtered_frame, text=value, font=("Arial", 11)).grid(
                row=idx, column=1, sticky=tk.W, pady=3
            )
        breakdown = metrics.get("status_breakdown", {})
        breakdown_frame = ttk.LabelFrame(content, text="Premium by Status", padding="10")
        breakdown_frame.pack(fill=tk.X)
        status_labels = {
            "prospects": "Prospects (no response yet)",
            "responded": "Responded", 
            "converted": "Converted Clients",
        }
        for idx, key in enumerate(["prospects", "responded", "converted"]):
            data = breakdown.get(key, {})
            ttk.Label(
                breakdown_frame,
                text=status_labels[key] + ":",
                font=("Arial", 11, "bold"),
            ).grid(row=idx, column=0, sticky=tk.W, pady=3)
            summary_text = (
                f"Count: {data.get('total_customers', 0)}  |  "
                f"Premium: ${data.get('total_premium', 0.0):,.2f}  |  "
                f"Response Rate: {data.get('response_rate', 0.0):.1f}%"
            )
            ttk.Label(
                breakdown_frame,
                text=summary_text,
                font=("Arial", 10),
            ).grid(row=idx, column=1, sticky=tk.W, pady=3)
    button_frame = ttk.Frame(container, padding="5")
    button_frame.grid(row=4, column=0, columnspan=5, sticky="ew", pady=(15, 0))
    for index in range(3):
        button_frame.columnconfigure(index, weight=1)

    ttk.Button(button_frame, text="Save Customer", command=save_selected_customer).grid(
        row=0, column=0, sticky="ew", padx=5
    )
    ttk.Button(button_frame, text="Clear", command=clear_form).grid(
        row=0, column=1, sticky="ew", padx=5
    )
    ttk.Button(button_frame, text="Refresh", command=refresh_tree).grid(
        row=0, column=2, sticky="ew", padx=5
    )

    tree.bind("<<TreeviewSelect>>", on_select)

    def on_close():
        nonlocal customer_rows, name_index, search_trace_id
        global customer_window
        customer_rows = []
        name_index = []
        if search_trace_id is not None:
            search_var.trace_remove("write", search_trace_id)
            search_trace_id = None
        clear_form()
        clear_suggestions()
        if customer_window is not None:
            window_to_close = customer_window
            customer_window = None
            window_to_close.destroy()

    customer_window.protocol("WM_DELETE_WINDOW", on_close)
    def perform_search(event=None):
        nonlocal current_search_matches
        query = search_var.get().strip()
        clear_suggestions()
        if not query:
            current_search_matches = None
            apply_filters()
            focus_first_result()
            return

        lowered = query.lower()
        scored_matches: List[tuple[int, Dict[str, object]]] = []
        for customer in customer_rows:
            name_value = str(customer.get("name", ""))
            email_value = str(customer.get("email", ""))
            phone_value = re.sub(r"\D", "", str(customer.get("phone", "")))
            scores = []
            for value in (name_value, email_value, phone_value):
                if not value:
                    continue
                comparison_source = value.lower()
                scores.append(fuzz.partial_ratio(lowered, comparison_source))
            best_score = max(scores, default=0)
            if best_score >= 60:
                scored_matches.append((best_score, customer))

        if not scored_matches:
            messagebox.showinfo(
                "No Matches",
                "No customers were found matching your search. Try a different name or spelling.",
            )
            current_search_matches = None
            apply_filters()
            return

        scored_matches.sort(key=lambda item: item[0], reverse=True)
        current_search_matches = [customer for _, customer in scored_matches]
        apply_filters()
        focus_first_result()

    search_entry.bind("<Return>", perform_search)

    def watch_search(*_):
        nonlocal current_search_matches
        if current_search_matches is not None:
            current_search_matches = None
        apply_filters()
        update_suggestion_box(search_var.get().strip())

    search_trace_id = search_var.trace_add("write", watch_search)

    refresh_tree()
    search_entry.focus_set()

# Signature selection
signature_label = tk.Label(main_frame, text="Signature:", font=("Arial", 12), bg="#f0f4f8")
signature_label.grid(row=1, column=0, sticky=tk.W, pady=5)
signature_choices = sorted(signature_profiles.keys())
default_signature = (
    "Brian Jones" if "Brian Jones" in signature_profiles else (signature_choices[0] if signature_choices else "")
)
signature_var = tk.StringVar(value=default_signature)
signature_dropdown = ttk.Combobox(
    main_frame,
    textvariable=signature_var,
    values=signature_choices,
    state="readonly",
)
signature_dropdown.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)

def update_signature_choices(selected=None):
    values = sorted(signature_profiles.keys())
    signature_dropdown.config(values=values)
    if not values:
        signature_var.set("")
        return
    desired = selected or signature_var.get()
    if desired not in values:
        desired = values[0]
    signature_var.set(desired)

def toggle_fullscreen(enabled: bool) -> None:
    try:
        root.attributes("-fullscreen", enabled)
    except tk.TclError:
        root.state("zoomed" if enabled else "normal")
    if not enabled:
        root.state("normal")


def show_about_dialog() -> None:
    about_text = dedent(
        """
        Auto Mailer Pro

        Built with the support of GPT-5-Codex to help automate multi-stage marketing campaigns for Jones Insurance Advisors.
        This distribution bundles a local database so each workstation tracks its own outreach and premium performance.
        """
    ).strip()
    messagebox.showinfo("About Auto Mailer Pro", about_text)


def show_instructions() -> None:
    instructions = dedent(
        """
        1. Choose your signature profile or add/remove a teammate from the File menu.
        2. Select the campaign mode, upload your sales data, and pick a ready-made letter template or author your own.
        3. Run the campaign to generate personalized letters, envelopes, and reports.
        4. Use Reports → Customer Database to search contacts, log responses, and build premium tracking summaries.

        Tip: The Reports window now includes instant search with auto-complete and checkboxes to filter totals by prospects,
        responses, and converted clients.
        """
    ).strip()
    messagebox.showinfo("Auto Mailer Pro Instructions", instructions)


def open_add_user_dialog() -> None:
    add_window = tk.Toplevel(root)
    add_window.title("Add Signature Profile")
    add_window.geometry("420x280")
    add_window.resizable(False, False)
    add_window.transient(root)
    add_window.grab_set()

    frame = ttk.Frame(add_window, padding=15)
    frame.pack(fill=tk.BOTH, expand=True)
    for column in range(2):
        frame.columnconfigure(column, weight=1)

    name_var = tk.StringVar()
    title_var = tk.StringVar()
    email_var = tk.StringVar()
    image_var = tk.StringVar()

    ttk.Label(frame, text="Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
    ttk.Entry(frame, textvariable=name_var).grid(row=0, column=1, sticky="ew", pady=5)

    ttk.Label(frame, text="Title:").grid(row=1, column=0, sticky=tk.W, pady=5)
    ttk.Entry(frame, textvariable=title_var).grid(row=1, column=1, sticky="ew", pady=5)

    ttk.Label(frame, text="Email:").grid(row=2, column=0, sticky=tk.W, pady=5)
    ttk.Entry(frame, textvariable=email_var).grid(row=2, column=1, sticky="ew", pady=5)

    ttk.Label(frame, text="Signature Image:").grid(row=3, column=0, sticky=tk.W, pady=5)
    ttk.Entry(frame, textvariable=image_var).grid(row=3, column=1, sticky="ew", pady=5)

    def browse_signature_image():
        file_path = filedialog.askopenfilename(
            title="Select Signature Image",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"),
                ("All files", "*.*"),
            ],
        )
        if file_path:
            image_var.set(file_path)

    ttk.Button(frame, text="Browse", command=browse_signature_image).grid(row=3, column=2, padx=5, pady=5)

    def save_new_signature():
        name = name_var.get().strip()
        if not name:
            messagebox.showerror("Validation Error", "Please provide the team member's name.")
            return
        if name in DEFAULT_SIGNATURE_PROFILES:
            messagebox.showerror(
                "Validation Error",
                "That name matches a built-in signature profile. Please choose a unique name.",
            )
            return

        title = title_var.get().strip()
        email = email_var.get().strip()
        image_path_input = image_var.get().strip()
        stored_image_path = None

        if image_path_input:
            source = Path(image_path_input)
            if not source.exists():
                messagebox.showerror("Validation Error", "The selected signature image could not be found.")
                return
            CUSTOM_SIGNATURES_DIR.mkdir(parents=True, exist_ok=True)
            suffix = source.suffix or ""
            base_name = sanitize_filename(name)
            destination = CUSTOM_SIGNATURES_DIR / f"{base_name}{suffix}"
            counter = 1
            while destination.exists():
                destination = CUSTOM_SIGNATURES_DIR / f"{base_name}_{counter}{suffix}"
                counter += 1
            try:
                shutil.copy2(source, destination)
            except OSError as exc:
                messagebox.showerror("Error", f"Unable to copy signature image: {exc}")
                return
            stored_image_path = destination

        signature_profiles[name] = (name, title, stored_image_path, email)
        persist_custom_signatures()
        update_signature_choices(name)
        messagebox.showinfo("Signature Saved", f"Signature profile for {name} has been added.")
        add_window.destroy()

    button_row = ttk.Frame(frame)
    button_row.grid(row=4, column=0, columnspan=3, pady=(15, 0))
    ttk.Button(button_row, text="Save", command=save_new_signature).grid(row=0, column=0, padx=5)
    ttk.Button(button_row, text="Cancel", command=add_window.destroy).grid(row=0, column=1, padx=5)
def open_remove_user_dialog() -> None:
    removable = sorted(
        name for name in signature_profiles if name not in DEFAULT_SIGNATURE_PROFILES
    )
    if not removable:
        messagebox.showinfo(
            "No Custom Signatures",
            "There are no custom signature profiles available to remove.",
        )
        return

    remove_window = tk.Toplevel(root)
    remove_window.title("Remove Signature Profile")
    remove_window.geometry("380x180")
    remove_window.resizable(False, False)
    remove_window.transient(root)
    remove_window.grab_set()

    frame = ttk.Frame(remove_window, padding=15)
    frame.pack(fill=tk.BOTH, expand=True)
    frame.columnconfigure(1, weight=1)

    ttk.Label(frame, text="Select Profile:").grid(row=0, column=0, sticky=tk.W, pady=5)
    selection_var = tk.StringVar(value=removable[0])
    ttk.Combobox(
        frame,
        textvariable=selection_var,
        values=removable,
        state="readonly",
    ).grid(row=0, column=1, columnspan=2, sticky="ew", pady=5)

    def confirm_removal() -> None:
        selected = selection_var.get()
        if not selected:
            return
        if not messagebox.askyesno(
            "Confirm Removal",
            f"Are you sure you want to delete the signature profile for {selected}?",
        ):
            return

        profile = signature_profiles.pop(selected, None)
        if not profile:
            messagebox.showerror(
                "Removal Error", f"Unable to locate the signature profile for {selected}."
            )
            return

        _, _, image_path, _ = profile
        if image_path:
            try:
                candidate = Path(image_path)
                try:
                    within_custom = candidate.is_relative_to(CUSTOM_SIGNATURES_DIR)
                except AttributeError:
                    within_custom = str(candidate).startswith(str(CUSTOM_SIGNATURES_DIR))
                if within_custom and candidate.exists():
                    candidate.unlink()
            except OSError as exc:
                messagebox.showwarning(
                    "File Warning",
                    f"The signature image could not be deleted: {exc}",
                )

        persist_custom_signatures()
        update_signature_choices()
        messagebox.showinfo(
            "Signature Removed", f"The signature profile for {selected} has been removed."
        )
        remove_window.destroy()

    button_row = ttk.Frame(frame)
    button_row.grid(row=1, column=0, columnspan=3, pady=(20, 0))
    ttk.Button(button_row, text="Remove", command=confirm_removal).grid(row=0, column=0, padx=5)
    ttk.Button(button_row, text="Cancel", command=remove_window.destroy).grid(
        row=0, column=1, padx=5
    )



update_signature_choices(default_signature)

menubar = tk.Menu(root)

file_menu = tk.Menu(menubar, tearoff=0)
file_menu.add_command(label="Add User…", command=open_add_user_dialog)
file_menu.add_command(label="Remove User…", command=open_remove_user_dialog)
file_menu.add_separator()
file_menu.add_command(label="Exit", command=root.destroy)
menubar.add_cascade(label="File", menu=file_menu)

reports_menu = tk.Menu(menubar, tearoff=0)
reports_menu.add_command(label="Customer Database", command=open_customer_manager)
menubar.add_cascade(label="Reports", menu=reports_menu)

view_menu = tk.Menu(menubar, tearoff=0)
view_menu.add_command(label="Enter Fullscreen", command=lambda: toggle_fullscreen(True))
view_menu.add_command(label="Exit Fullscreen", command=lambda: toggle_fullscreen(False))
menubar.add_cascade(label="View", menu=view_menu)

about_menu = tk.Menu(menubar, tearoff=0)
about_menu.add_command(label="About", command=show_about_dialog)
about_menu.add_command(label="Instructions", command=show_instructions)
menubar.add_cascade(label="About", menu=about_menu)

root.config(menu=menubar)


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


# Run button
run_button = ttk.Button(main_frame, text="Run Campaign", command=run_campaign, style="TButton")
run_button.grid(row=9, column=0, columnspan=4, pady=20)
run_button.grid(row=8, column=0, columnspan=4, pady=20)

# Progress bar
progress_bar = ttk.Progressbar(main_frame, mode='indeterminate')
progress_bar.grid(row=10, column=0, columnspan=4, sticky=(tk.W, tk.E), pady=5)
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

if __name__ == "__main__":
    root.mainloop()