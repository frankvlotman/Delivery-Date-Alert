import win32com.client as win32
from datetime import datetime, timedelta

# Dummy purchase order data (replace with your actual data)
purchase_orders = {
    'PO001': {'vendor': 'Vendor A', 'delivery_date': datetime(2024, 12, 31)},
    'PO002': {'vendor': 'Vendor B', 'delivery_date': datetime(2024, 6, 20)},
    'PO003': {'vendor': 'Vendor C', 'delivery_date': datetime(2024, 6, 25)},
    'PO004': {'vendor': 'Vendor D', 'delivery_date': datetime(2024, 7, 1)},
}

def send_email(upcoming_orders):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'example@example.com'
    mail.Subject = "Upcoming Delivery Date Notification"
    mail.Body = "The following purchase orders have upcoming delivery dates within 5 days or have already passed:\n\n"
    for po, details in upcoming_orders.items():
        mail.Body += f"PO: {po}\nVendor: {details['vendor']}\nDelivery Date: {details['delivery_date'].strftime('%Y-%m-%d')}\n\n"
    mail.Send()

def check_upcoming_delivery():
    upcoming_orders = {}
    current_date = datetime.now()
    for po, details in purchase_orders.items():
        days_difference = (details['delivery_date'] - current_date).days
        if days_difference <= 5 and days_difference >= 0 or details['delivery_date'] < current_date:
            upcoming_orders[po] = details
    if upcoming_orders:
        send_email(upcoming_orders)
        print("Email notification sent for upcoming delivery dates or past due dates.")
    else:
        print("No upcoming or past due delivery dates.")

if __name__ == "__main__":
    check_upcoming_delivery()
