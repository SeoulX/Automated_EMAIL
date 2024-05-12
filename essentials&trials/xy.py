import pyautogui
import pygetwindow

# Get a list of all open windows
all_windows = pygetwindow.getAllWindows()

# Print the title of each window
for window in all_windows:
    print(f"Window Title: {window.title}")
    
gmail_url = 'https://mail.google.com/'
gmail_window = None
for window in pyautogui.getAllWindows():
    if gmail_url in window.title.lower() or 'aandrianbinas01@gmail.com - gmail' in window.title.lower():
        gmail_window = window
        break

if gmail_window:
    gmail_window.activate()
