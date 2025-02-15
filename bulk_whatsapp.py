import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import pandas as pd
import time
import threading
import pyautogui
import os
import phonenumbers

# Function to send WhatsApp messages using the desktop application
def send_messages():
    try:
        # Disable the Send button to prevent multiple executions
        send_button.config(state=tk.DISABLED)
        stop_button.config(state=tk.NORMAL)

        # Get user inputs
        message = message_box.get("1.0", tk.END).strip()
        interval = int(interval_box.get())
        attachment = attachment_path.get()

        # Validate inputs
        if not message:
            messagebox.showerror("Error", "Please enter a message.")
            send_button.config(state=tk.NORMAL)
            stop_button.config(state=tk.DISABLED)
            return

        if not phone_numbers:
            messagebox.showerror("Error", "Please upload a valid contact list.")
            send_button.config(state=tk.NORMAL)
            stop_button.config(state=tk.DISABLED)
            return

        # Confirm sending
        confirmation = messagebox.askyesno("Confirmation", "Are you sure you want to send the messages?")
        if not confirmation:
            send_button.config(state=tk.NORMAL)
            stop_button.config(state=tk.DISABLED)
            return

        # Start sending messages
        status_label.config(text="Sending messages...", fg="blue")
        app.update()

        for i in range(last_sent_index[0], len(phone_numbers)):
            if stop_event.is_set():
                status_label.config(text="Message sending stopped.", fg="red")
                break
            try:
                number, country = phone_numbers[i]
                number = number.lstrip('+')  # Remove one '+' symbol
                # Open WhatsApp Desktop and search for the contact
                pyautogui.hotkey('ctrl', 'alt', 'w')  # Shortcut to open WhatsApp Desktop (customize if needed)
                time.sleep(2)
                pyautogui.click(200, 100)  # Click on the search bar (adjust coordinates for your screen)
                time.sleep(1)
                pyautogui.typewrite(number)
                time.sleep(2)
                pyautogui.press('enter')

                # Type and send the message
                pyautogui.typewrite(message)
                time.sleep(1)
                pyautogui.press('enter')

                # Attach file if provided
                if attachment:
                    pyautogui.click(100, 700)  # Click on the attachment icon (adjust coordinates for your screen)
                    time.sleep(2)
                    pyautogui.typewrite(attachment)
                    time.sleep(1)
                    pyautogui.press('enter')

                sent_log.append((number, "Sent", country))
                last_sent_index[0] = i + 1  # Save the last sent index
                update_number_list()
                time.sleep(interval)  # Delay between messages
                status_label.config(text=f"Sent {i + 1}/{len(phone_numbers)} messages.")
                app.update()
            except Exception as e:
                error_log.append(f"Failed to send to {number}: {str(e)}")
                sent_log.append((number, "Error", country))

        if not stop_event.is_set():
            status_label.config(text="All messages sent successfully!", fg="green")
        if error_log:
            with open("error_log.txt", "w") as f:
                f.write("\n".join(error_log))
            messagebox.showinfo("Completed with Errors", "Some messages failed to send. Check 'error_log.txt' for details.")

        stop_event.clear()
        send_button.config(state=tk.NORMAL)
        stop_button.config(state=tk.DISABLED)
    except Exception as e:
        messagebox.showerror("Error", str(e))
        send_button.config(state=tk.NORMAL)
        stop_button.config(state=tk.DISABLED)

# Function to upload contact list
def upload_contacts():
    global phone_numbers
    filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv")])
    if filepath:
        try:
            if filepath.endswith(".xlsx"):
                df = pd.read_excel(filepath)
            elif filepath.endswith(".csv"):
                df = pd.read_csv(filepath)
            else:
                messagebox.showerror("Error", "Unsupported file format.")
                return

            if "Phone Number" in df.columns:
                phone_numbers = []
                for num in df["Phone Number"]:
                    if pd.notna(num):
                        try:
                            parsed_number = phonenumbers.parse(f"+{str(num).strip()}")
                            country_name = phonenumbers.region_code_for_number(parsed_number)
                            phone_numbers.append((f"+{str(num).strip()}", country_name))
                        except phonenumbers.NumberParseException:
                            phone_numbers.append((f"+{str(num).strip()}", "Invalid"))

                last_sent_index[0] = 0
                update_number_list()
                messagebox.showinfo("Success", f"Loaded {len(phone_numbers)} contacts.")
            else:
                messagebox.showerror("Error", "The file must contain a 'Phone Number' column.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load contacts: {str(e)}")

# Function to select attachment file
def select_attachment():
    filepath = filedialog.askopenfilename()
    if filepath:
        attachment_path.set(filepath)

# Function to save the log of sent messages
def save_sent_log():
    try:
        if sent_log:
            with open("sent_log.txt", "w") as f:
                for number, status, country in sent_log:
                    f.write(f"{number},{status},{country}\n")
            messagebox.showinfo("Log Saved", "Sent messages log saved to 'sent_log.txt'.")
        else:
            messagebox.showinfo("No Log", "No messages have been sent yet.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save log: {str(e)}")

# Function to stop the message sending process
def stop_sending():
    stop_event.set()

# Function to update the number list in the GUI
def update_number_list():
    number_list.delete(0, tk.END)
    for i, (number, country) in enumerate(phone_numbers):
        status = "Sent" if i < last_sent_index[0] else "Pending"
        number_list.insert(tk.END, f"{number} ({country}) - {status}")

# Initialize the main GUI application
app = tk.Tk()
app.title("Advanced WhatsApp Bulk Messaging")
app.geometry("800x700")
app.configure(bg="#f0f0f0")
app.resizable(True, True)  # Allow window to be maximized and minimized

# Global variables
phone_numbers = []
attachment_path = tk.StringVar()
error_log = []
sent_log = []
stop_event = threading.Event()
last_sent_index = [0]  # To track the index of the last sent message

# GUI Layout with enhanced design
tk.Label(app, text="Message:", font=("Arial", 12, "bold"), bg="#f0f0f0").pack(pady=5)
message_box = tk.Text(app, height=10, width=60, font=("Arial", 10))  # Increased height to 10
message_box.pack(pady=5)

frame = tk.Frame(app, bg="#f0f0f0")
frame.pack(pady=10)

tk.Label(frame, text="Interval (seconds):", font=("Arial", 10), bg="#f0f0f0").grid(row=0, column=0, padx=5)
interval_box = tk.Entry(frame, font=("Arial", 10), width=10)
interval_box.grid(row=0, column=1, padx=5)
interval_box.insert(0, "25")

tk.Button(frame, text="Upload Contact List", command=upload_contacts, bg="#0078d7", fg="white", font=("Arial", 10)).grid(row=0, column=2, padx=10)

attachment_frame = tk.Frame(app, bg="#f0f0f0")
attachment_frame.pack(pady=10)

tk.Button(attachment_frame, text="Attach File", command=select_attachment, bg="#0078d7", fg="white", font=("Arial", 10)).grid(row=0, column=0, padx=10)

attachment_entry = tk.Entry(attachment_frame, textvariable=attachment_path, state="readonly", width=40, font=("Arial", 10))
attachment_entry.grid(row=0, column=1, padx=10)

send_button = tk.Button(app, text="Send Messages", command=lambda: threading.Thread(target=send_messages).start(), bg="green", fg="white", font=("Arial", 12, "bold"))
send_button.pack(pady=10)

stop_button = tk.Button(app, text="Stop Sending", command=stop_sending, bg="red", fg="white", font=("Arial", 12, "bold"), state=tk.DISABLED)
stop_button.pack(pady=5)

tk.Button(app, text="Save Sent Log", command=save_sent_log, bg="#0078d7", fg="white", font=("Arial", 10)).pack(pady=10)

status_label = tk.Label(app, text="", font=("Arial", 10), bg="#f0f0f0")
status_label.pack(pady=5)

number_list = tk.Listbox(app, font=("Arial", 10), width=100, height=15)
number_list.pack(pady=10)

# Run the application
app.mainloop()