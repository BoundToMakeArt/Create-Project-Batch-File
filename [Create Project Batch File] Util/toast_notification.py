
from win10toast import ToastNotifier

def show_notification(title, message, icon_path=None):
    toast = ToastNotifier()
    toast.show_toast(title, message, icon_path=icon_path, duration=10)

if __name__ == "__main__":
    # Change these values as needed
    notification_title = "Everything you selected is now open"
    notification_message = "Good luck with your project!"
    notification_icon_path = "your_path\\icon.ico"  # Dynamic path to icon.ico

    show_notification(notification_title, notification_message, notification_icon_path)
