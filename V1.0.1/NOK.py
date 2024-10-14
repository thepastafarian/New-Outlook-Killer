import time
import subprocess
import psutil
import pystray
from PIL import Image, ImageDraw
from threading import Thread
import tkinter as tk
from tkinter import ttk

# PowerShell command to uninstall Outlook
uninstall_command = 'powershell "Get-AppxPackage *Outlook* | Remove-AppxPackage"'

# Mail app process name (this may vary, adjust if needed)
mail_app_process = "HxOutlook.exe"
scan_interval = 5  # Default scan interval in seconds
resource_usage = {"cpu": 0, "memory": 0}

# Function to run the PowerShell uninstall command
def uninstall_outlook():
    try:
        print("Uninstalling Outlook...")
        result = subprocess.run(uninstall_command, shell=True, check=True, capture_output=True, text=True)
        print(f"Uninstall output: {result.stdout}")
    except subprocess.CalledProcessError as e:
        print(f"Error during uninstall: {e}")
        print(f"Uninstall error output: {e.stderr}")

# Function to monitor if the Mail app is opened or closed
def monitor_mail_app():
    global resource_usage
    mail_running = False
    print("Monitoring for Mail app...")

    while True:
        # Check if the Mail app is running
        mail_running_now = any(proc.info['name'] == mail_app_process for proc in psutil.process_iter(['name']))

        # If the Mail app is launched
        if mail_running_now and not mail_running:
            print("Mail app launched. Running Outlook uninstall command...")
            uninstall_outlook()  # Run the PowerShell script to uninstall Outlook
            mail_running = True  # Update the state to reflect the Mail app is running

        # If the Mail app is running, uninstall Outlook every scan interval
        if mail_running_now:
            print("Mail app is running. Uninstalling Outlook again...")
            uninstall_outlook()  # Run the PowerShell script every scan interval

        # If the Mail app was running but is now closed
        if not mail_running_now and mail_running:
            print("Mail app closed. Running Outlook uninstall command...")
            uninstall_outlook()  # Run the PowerShell script to uninstall Outlook when the Mail app closes
            mail_running = False  # Update the state to reflect the Mail app is closed

        # Capture CPU and memory usage
        process = psutil.Process()
        resource_usage["cpu"] = process.cpu_percent(interval=0.1)
        resource_usage["memory"] = process.memory_info().rss / (1024 * 1024)  # Convert to MB

        time.sleep(scan_interval)  # Check every scan interval

# Function to create an image for the tray icon
def create_image(width, height, color1, color2):
    image = Image.new('RGB', (width, height), color1)
    dc = ImageDraw.Draw(image)
    dc.rectangle((width // 2, 0, width, height // 2), fill=color2)
    dc.rectangle((0, height // 2, width // 2, height), fill=color2)
    return image

# Function to quit the system tray app
def quit_app(icon, item):
    icon.stop()

# Function to show the status window
def show_status_window():
    global scan_interval, resource_usage

    # Create a window
    window = tk.Tk()
    window.title("NOK Monitor")
    window.geometry("300x200")

    # Mail app status label
    mail_status_label = ttk.Label(window, text="Mail App Status: Unknown")
    mail_status_label.pack(pady=5)

    # Scan interval adjustment
    interval_label = ttk.Label(window, text="Set Scan Interval (seconds):")
    interval_label.pack(pady=5)
    interval_entry = ttk.Entry(window)
    interval_entry.pack(pady=5)
    interval_entry.insert(0, str(scan_interval))

    # Update interval button
    def update_interval():
        global scan_interval
        try:
            scan_interval = int(interval_entry.get())
        except ValueError:
            pass  # Ignore invalid input
    update_button = ttk.Button(window, text="Update Interval", command=update_interval)
    update_button.pack(pady=5)

    # Resource usage display
    cpu_label = ttk.Label(window, text="CPU Usage: 0%")
    cpu_label.pack(pady=5)
    mem_label = ttk.Label(window, text="Memory Usage: 0 MB")
    mem_label.pack(pady=5)

    # Function to update the status
    def update_status():
        # Check Mail app status
        mail_running_now = any(proc.info['name'] == mail_app_process for proc in psutil.process_iter(['name']))
        mail_status_label.config(text=f"Mail App Status: {'Running' if mail_running_now else 'Not Running'}")

        # Update resource usage labels
        cpu_label.config(text=f"CPU Usage: {resource_usage['cpu']:.2f}%")
        mem_label.config(text=f"Memory Usage: {resource_usage['memory']:.2f} MB")

        # Schedule the next update
        window.after(1000, update_status)

    # Start updating status
    update_status()

    # Run the window loop
    window.mainloop()

# Function to run the tray icon
def run_tray_icon():
    # Create the tray icon
    icon_image = create_image(64, 64, 'black', 'green')
    icon = pystray.Icon("NOK", icon_image, "NOK.exe", menu=pystray.Menu(
        pystray.MenuItem("Show Status", lambda: Thread(target=show_status_window).start()),
        pystray.MenuItem("Quit", quit_app)
    ))
    icon.run()

if __name__ == "__main__":
    # Start the tray icon in a separate thread
    tray_thread = Thread(target=run_tray_icon)
    tray_thread.start()

    # Uninstall Outlook immediately on system startup
    uninstall_outlook()

    # Start monitoring for the Mail app
    monitor_mail_app()
