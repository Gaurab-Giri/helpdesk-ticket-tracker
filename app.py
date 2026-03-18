from db import get_connection, setup_database
from datetime import datetime

def create_ticket(title, category, priority):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "INSERT INTO tickets (title, category, priority) VALUES (%s, %s, %s)",
        (title, category, priority)
    )
    conn.commit()
    print(f"Ticket created: {title}")
    cursor.close()
    conn.close()

def view_tickets(status=None):
    conn = get_connection()
    cursor = conn.cursor()
    if status:
        cursor.execute("SELECT * FROM tickets WHERE status = %s", (status,))
    else:
        cursor.execute("SELECT * FROM tickets")
    tickets = cursor.fetchall()
    if not tickets:
        print("No tickets found.")
        return
    print(f"\n{'ID':<5} {'Title':<30} {'Category':<15} {'Priority':<10} {'Status':<10}")
    print("-" * 75)
    for t in tickets:
        print(f"{t[0]:<5} {t[1]:<30} {t[2]:<15} {t[3]:<10} {t[4]:<10}")
    cursor.close()
    conn.close()

def resolve_ticket(ticket_id):
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute(
        "UPDATE tickets SET status='Resolved', resolved_at=%s WHERE id=%s",
        (datetime.now(), ticket_id)
    )
    conn.commit()
    print(f"Ticket {ticket_id} marked as resolved.")
    cursor.close()
    conn.close()

def export_to_excel():
    import openpyxl
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM tickets")
    tickets = cursor.fetchall()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tickets"
    ws.append(["ID", "Title", "Category", "Priority", "Status", "Created", "Resolved"])
    for t in tickets:
        ws.append(list(t))
    wb.save("tickets_report.xlsx")
    print("Exported to tickets_report.xlsx")
    cursor.close()
    conn.close()

def menu():
    while True:
        print("\n===== IT Help Desk Tracker =====")
        print("1. Create new ticket")
        print("2. View all tickets")
        print("3. View open tickets only")
        print("4. Resolve a ticket")
        print("5. Export to Excel")
        print("6. Exit")
        choice = input("Choose an option (1-6): ").strip()

        if choice == "1":
            title = input("Issue title: ").strip()
            print("Categories: Network, Software, Endpoint, Hardware, Other")
            category = input("Category: ").strip()
            print("Priorities: High, Medium, Low")
            priority = input("Priority: ").strip()
            create_ticket(title, category, priority)

        elif choice == "2":
            view_tickets()

        elif choice == "3":
            view_tickets(status="Open")

        elif choice == "4":
            view_tickets()
            ticket_id = input("\nEnter ticket ID to resolve: ").strip()
            resolve_ticket(int(ticket_id))

        elif choice == "5":
            export_to_excel()

        elif choice == "6":
            print("Goodbye!")
            break

        else:
            print("Invalid option. Please choose 1-6.")

if __name__ == "__main__":
    setup_database()
    menu()
