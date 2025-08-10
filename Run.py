import tkinter as tk
from tkinter import scrolledtext
import sys
import threading
from AutoMailerPro_v5_1 import main

FONT = ("Segoe UI", 11)
BUTTON_BG = "#0078D7"
BUTTON_FG = "white"
BG_COLOR = "#F0F0F0"
TEXT_BG = "white"
TEXT_FG = "black"

class StdoutRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, s):
        self.text_widget.config(state='normal')
        self.text_widget.insert(tk.END, s)
        self.text_widget.see(tk.END)
        self.text_widget.config(state='disabled')

    def flush(self):
        pass

def run_script():
    run_button.config(state='disabled')
    status_var.set("Running script...")
    threading.Thread(target=threaded_main).start()

def threaded_main():
    try:
        main()
    except Exception as e:
        print(f"Error: {e}")
    run_button.config(state='normal')
    status_var.set("Ready")

root = tk.Tk()
root.title("InsuranceMailer GUI")
root.geometry("700x500")
root.configure(bg=BG_COLOR)

frame = tk.Frame(root, bg=BG_COLOR, padx=15, pady=15)
frame.pack(fill='both', expand=True)
credits_text = (
    "AutoMailerPro v5.1\n"
    "Author: Kyle Padilla\n"
    "Last Updated: 08/09/2025\n"
    "Jones Insurance Advisors, Inc."
)

credits_label = tk.Label(frame, text=credits_text, bg=BG_COLOR, font=("Segoe UI", 9), justify="center")
credits_label.pack(pady=(0, 15))


run_button = tk.Button(frame, text="Run AutoMailerPro_v5.1", command=run_script,
                       font=FONT, bg=BUTTON_BG, fg=BUTTON_FG,
                       activebackground="#005a9e", activeforeground="white",
                       relief="flat", padx=10, pady=5)
run_button.pack(pady=(0, 15))
boat_art = """
                  __/___
         _____/______|
 _______/_____\_______\_____
 \              < < <       |
  ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
"""

boat_label = tk.Label(frame, text=boat_art, bg=BG_COLOR, font=("Courier New", 12), justify="center")
boat_label.pack(pady=(5,15))

output_text = scrolledtext.ScrolledText(frame, width=80, height=20, state='disabled',
                                        font=("Consolas", 10), bg=TEXT_BG, fg=TEXT_FG,
                                        bd=2, relief="sunken", insertbackground="black")
output_text.pack(fill='both', expand=True)

sys.stdout = StdoutRedirector(output_text)
sys.stderr = StdoutRedirector(output_text)

status_var = tk.StringVar()
status_var.set("Ready")
status_bar = tk.Label(root, textvariable=status_var, bd=1, relief='sunken',
                      anchor='w', font=("Segoe UI", 9))
status_bar.pack(side='bottom', fill='x')

root.mainloop()
