import os
import time
import pandas as pd
import pyautogui
import subprocess
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

abort_flag = False  # Global flag to control the sending loop

def extract_phone_number(phone):
    phone = re.sub(r'^\+\d{1,2}', '', phone)
    return phone


def send_whatsapp(data_file_excel):
    global abort_flag
    # Read the Excel file
    df = pd.read_excel(data_file_excel, dtype={"Contact": str})
    names = df["Name"].values
    contacts = df["Contact"].values
    # messages = df['Message'].values

    # Open WhatsApp Desktop application using URI scheme
    subprocess.Popen(['start', 'whatsapp://'], shell=True)
    time.sleep(3)  # Allow time for WhatsApp to open

    # for name, contact, message in zip(names, contacts, messages):
    for name, contact in zip(names, contacts):
        if abort_flag:  # Check the abort flag
            print("Aborted!")
            messagebox.showinfo("Aborted", "Message sending process aborted!")
            break

        contact = extract_phone_number(contact)
        message = "🔹 Tescom Business Internet Service 🔹 \n\nBoost your business with our reliable high-speed internet! \n\n📶 Plans:  \n\n50 Mbps - ₹600/month \n100 Mbps - ₹900/month\n150 Mbps - ₹1100/month\n✨ Installation & Equipment:\n\nInstallation Fee: ₹1500 (one-time)\nRouter: ₹1000 (one-time)\n🎉 Special Offer: Free Wi-Fi Router with any yearly plan!\n\n📞 Contact Us:\nMessage or call 7738999055 for details and sign-up.\n\nStay connected with Tescom Business Internet Service!"
        
        time.sleep(1)
        pyautogui.click(x=445, y=87)
        time.sleep(0.7)
        pyautogui.typewrite(contact)   # write/find/search contact number
        time.sleep(0.2)
        pyautogui.hotkey('tab')
        time.sleep(0.1)
        pyautogui.press('enter')   
        time.sleep(0.1)
        pyautogui.click(x=701, y=987)
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(1.5)
        pyautogui.press('enter')   

        print(f"Message sent to {name} ({contact})")
    print("Done!")

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        excel_path.set(file_path)
        status_label.config(text=f"Selected file: {file_path}")

def run_messaging():
    global abort_flag
    abort_flag = False  # Reset abort flag before starting
    file_path = excel_path.get()
    if file_path:
        try:
            send_whatsapp(file_path)
            print("file_path => ",file_path)
            if not abort_flag:
                messagebox.showinfo("Success", "Messages sent successfully!")
        except Exception as e:
            messagebox.showerror("Error", str(e))
    else:
        messagebox.showwarning("Warning", "Please select an Excel file first.")

def abort_messaging():
    global abort_flag
    abort_flag = True  # Set abort flag to stop the loop





# ________________________________________________________________________________________
# ________________________________________________________________________________________
# ________________________________Main Code start From here_______________________________
# ________________________________________________________________________________________
# ________________________________________________________________________________________

# Create the main application window
root = tk.Tk()
root.title("WhatsApp Message Sender")
root.state('zoomed')  # Maximize the screen by default

# Define styles
style = ttk.Style()
style.configure('TButton', font=('Helvetica', 12))

excel_path = tk.StringVar()

# Create and place widgets in the window
header = tk.Label(root, text="Bulk WhatsApp Message Sender", font=("Helvetica", 20, "bold"))
header.pack(pady=20)

frame = ttk.Frame(root, padding="10 10 10 10")
frame.pack(expand=True)

select_button = ttk.Button(frame, text="Select Excel File", command=select_file)
select_button.pack(pady=10, ipadx=20, ipady=10)

run_button = ttk.Button(frame, text="Send Messages", command=run_messaging)
run_button.pack(pady=10, ipadx=20, ipady=10)

abort_button = ttk.Button(frame, text="Abort", command=abort_messaging)
abort_button.pack(pady=10, ipadx=20, ipady=10)

status_label = ttk.Label(frame, text="No file selected", anchor="center")
status_label.pack(pady=20)


# footer_label = tk.Label(root, text="Developed by Pranav Sangave. For more software inquiries reach out us at +91 9096553454, sangways.web@gmail.com. (Android Apps | Websites | Desktop Apps)", font=("Arial", 8))
# footer_label.pack(side="bottom", pady=20)

# Start the Tkinter event loop
root.mainloop()
