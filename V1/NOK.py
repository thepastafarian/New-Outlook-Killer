import time
import subprocess
import psutil
import pystray
from PIL import Image, ImageDraw
from threading import Thread

# PowerShell command to uninstall Outlook
uninstall_command = 'powershell "Get-AppxPackage *Outlook* | Remove-AppxPackage"'

# Mail app process name (this may vary, adjust if needed)
mail_app_process = "HxOutlook.exe"

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

        # If the Mail app is running, uninstall Outlook every 5 seconds
        if mail_running_now:
            print("Mail app is running. Uninstalling Outlook again...")
            uninstall_outlook()  # Run the PowerShell script every 5 seconds

        # If the Mail app was running but is now closed
        if not mail_running_now and mail_running:
            print("Mail app closed. Running Outlook uninstall command...")
            uninstall_outlook()  # Run the PowerShell script to uninstall Outlook when the Mail app closes
            mail_running = False  # Update the state to reflect the Mail app is closed

        time.sleep(5)  # Check every 5 seconds

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

# Function to run the tray icon
def run_tray_icon():
    # Create the tray icon
    icon_image = create_image(64, 64, 'black', 'green')
    icon = pystray.Icon("Outlook Killer", icon_image, menu=pystray.Menu(
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
