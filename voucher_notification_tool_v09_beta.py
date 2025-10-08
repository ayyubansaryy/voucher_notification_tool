import os
import time
from io import StringIO
import datetime
import tkinter as tk
import pandas as pd
from tkcalendar import Calendar
from reportlab.lib.pagesizes import A4
import re

#### Set up messages
def boot_msg():
    print("\n══════════════════════════════════════════════════════\nVoucher Notification Text Tool\n══════════════════════════════════════════════════════\nv0.7 (developed by: FEL-89242)\nDefault output location: Desktop\n══════════════════════════════════════════════════════\n\n↻ AUTO LOADING DATA.....(please wait)\n\n")

def restart_msg():
    os.system("cls" if os.name == "nt" else "clear")
    print("\n🟢 Successfully Restarted\n\nPaste data below (with headers) & press 'Enter' twice:\n══════════════════════════════════════════════════════\n")

def select_date_msg():
    os.system("cls" if os.name == "nt" else "clear")
    print("\n🟢 Select date again to filter by...\n")

def closing():
    print("\n👋 Closing in 5 seconds...")
    time.sleep(3)


#### Pick validity dates from calendar
def pick_date(title="Select a date") -> str:
    root = tk.Tk()
    root.title(title)
    tk.Label(root, text=title, font=("Arial", 16, "bold")).pack(pady=10)
    cal = Calendar(root, selectmode="day", date_pattern="dd/mm/yyyy", font=("Arial", 14))
    cal.pack(pady=10)
    w, h = 500, 400
    x = (root.winfo_screenwidth() - w) // 2
    y = (root.winfo_screenheight() - h) // 2
    root.geometry(f"{w}x{h}+{x}+{y}")
    selected_date = tk.StringVar()
    def confirm():
        selected_date.set(cal.get_date())
        root.destroy()
    tk.Button(root, text="Confirm", font=("Arial", 12, "bold"), padx=20, pady=10,
              command=confirm).pack(pady=10)
    root.lift()
    root.attributes("-topmost", True)
    root.after(100, lambda: root.attributes("-topmost", False))
    root.mainloop()
    return selected_date.get()


#### Get validity dates
def get_validity_dates():
    start = format_date(pick_date("Validity Starts:"))
    end = format_date(pick_date("Validity Ends:"))
    return start, end


#### Format date to string
def format_date(date_input: str) -> str:
    try:
        return datetime.datetime.strptime(date_input, "%d/%m/%Y").strftime("%d %B")
    except ValueError:
        return date_input
    

#### Parse data from google sheet
def parse_data():
    sheet_id = "1oHZKJaLpJdbyuOd9A5RPd0tFw6sPq3tY28euOEX5urg"
    gid = "0"
    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid={gid}"
    
    try:
        df = pd.read_csv(url)
    except Exception as e:
        print(f"🛑 Failed to fetch dtat from Google Sheet: {e}")
        return None

    # Assume first column is date
    date_col = df.columns[0]
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)

    print("🔘 Select a valid date to filter by (from the pop-up calendar)\n")

    while True:
        filter_by = pick_date()
        try:
            user_date = datetime.datetime.strptime(filter_by, "%d/%m/%Y").date()
        except ValueError:
            print("🛑 Invalid date selected!\n")
            choice = input("🔘 Press ENTER ⏎ to Restart / Close window to Exit\n").strip()
            if choice == "":
                continue
            else:
                exit()

        # Filter by day and month only
        filtered_df = df[df[date_col].apply(lambda x: (x.day, x.month) if pd.notnull(x) else (0,0)) == (user_date.day, user_date.month)]

        if filtered_df.empty:
            print(f"🛑 No matching records found for '{user_date.strftime('%d %B')}'\n")
            choice = input("🔘 Press ENTER ⏎ to Restart / Close window to Exit\n").strip()
            if choice == "":
                select_date_msg()
                continue
            else:
                exit()

        # Handle 'Voucher Given' column
        if "Voucher Given" in filtered_df.columns:
            # Clean the column first
            col = (
                filtered_df["Voucher Given"]
                .astype(str)
                .str.replace("\u00A0", "", regex=False)
                .str.replace("\ufeff", "", regex=False)
                .str.strip()
                .str.lower()
            )

            # Fill every cell with "no" if it is not "yes" or "withdrawn"
            col = col.apply(lambda x: x if x in ["yes", "withdrawn"] else "no")

            # Assign back the cleaned column to the DataFrame
            filtered_df["Voucher Given"] = col

            # Now filter out rows where voucher is "yes" or "withdrawn"
            voucher_filtered = filtered_df[~col.isin(["yes", "withdrawn"])]

            print("🧩 Cleaned 'Voucher Given' unique values:", col.unique())
            print(f"✅ Rows kept for {user_date.strftime('%d %B')}: {len(voucher_filtered)}")

            if voucher_filtered.empty:
                print(f"🛑 All vouchers have already been issued for '{user_date.strftime('%d %B')}'\n")
                choice = input("🔘 Press ENTER ⏎ to Restart / Close window to Exit\n").strip()
                if choice == "":
                    select_date_msg()
                    continue
                else:
                    exit()

            return voucher_filtered

        return filtered_df


#### Build notification text segments
def build_segments(df: pd.DataFrame, start: str, end: str) -> list[str]:
    segments = []

    def format_order_contact(line: str) -> str:
        match = re.match(r"(\w+)\s*(\d+)", line)
        if match:
            order, contact = match.groups()
            return f"{order}\u00A0\u00A0\u00A0{contact}"
        return line\
    
    #0.1 Split date into two segmenst (e.g. 20 january --> ["20", "January"])
    start_a = start.split()[0].lstrip("0")
    start_b = start.split()[1].lstrip("0")
    end_a = end.split()[0].lstrip("0")
    end_b = end.split()[1].lstrip("0")
    #0.2 prepare start day extension
    if start_a in ("1", "31", "21"): start_date = f"{start_a}st {start_b}"
    elif start_a in ("2", "22"): start_date = f"{start_a}nd {start_b}"
    elif start_a in ("3", "23"): start_date = f"{start_a}rd {start_b}"
    else: start_date = f"{start_a}th {start_b}"
    #0.3 prepare end day extension
    if end_a in ("1", "31", "21"): end_date = f"{end_a}st {end_b}"
    elif end_a in ("2", "22") : end_date = f"{end_a}nd {end_b}"
    elif end_a in ("3", "23"): end_date = f"{end_a}rd {end_b}"
    else: end_date = f"{end_a}th {end_b}"

    # Build main notification text that users will receive
    for serial, (amount, group) in enumerate(df.groupby("Voucher"), start=1):
        code_str = f"SORRY{int(amount)}"
        mov = int(amount) + 49
        order_contact_lines = [
            format_order_contact(f"{row['Order No']} {row['Contact']}")
            for _, row in group.iterrows()
        ]
        lines = [
            f"{serial}. {code_str}",
            *order_contact_lines,
            f"Use coupon {code_str} to get {int(amount)} taka off",
            f"Minimum order: {mov} taka",
            f"Validity: {start_date} to {end_date}",
            "Not applicable for Flat discount-providing restaurants\n",
        ]
        segments.append("\n".join(lines))
    return segments


#### - - - - - - - - - - - - - - - - - Main Program Begins
boot_msg()
def main():
    while True:
        df = parse_data()
        print(df)
        if df is None:
            print("🛑 ERROR! data not found!\n\n")
            try:
                choice = input("🔘 Press ENTER ⏎ to Reload Data \n\n🔘 Close window to Exit\n\n").strip()
            except EOFError:
                print("\n👋 Window closed. Exiting tool...")
                exit()
            if choice == "":
                restart_msg()
                continue
            else:
                restart_msg()
                continue

        # Validate required columns
        required_cols = {"Order No", "Contact", "Voucher"}
        if not required_cols.issubset(df.columns):
            print(f"🛑 ERROR! Missing columns. Required: {', '.join(required_cols)}")
            restart_msg()
            continue

        # Check for missing voucher values
        missing_voucher_rows = df[df["Voucher"].isna()]
        if not missing_voucher_rows.empty:
            print("\n🛑 ERROR! Following rows have missing voucher amount:\n")
            for _, row in missing_voucher_rows.iterrows():
                print(f"{row['Order No']}  0{row['Contact']}  {row['Voucher']}")
            
            try:
                choice = input("🔘 Press ENTER ⏎ to Restart \n\n🔘 Close window to Exit\n\n").strip()
            except EOFError:
                print("\n👋 Window closed. Exiting tool...")
                exit()
            if choice == "":
                restart_msg()
                continue
            else:
                restart_msg()
                continue

        # Check for duplicate contacts
        duplicate_contacts = df[df.duplicated(subset=["Contact"], keep=False)]
        if not duplicate_contacts.empty:
            duplicate_contacts = duplicate_contacts.sort_values(by="Contact")
            print("\nDuplicate contacts found:\n")
            for _, row in duplicate_contacts.iterrows():
                print(f"{row['Order No']} 0{row['Contact']} {int(row['Voucher'])}")
            
            try:
                choice = input("\n🔘 Type and enter '𝐂' to continue \n\n🔘 Press ENTER ⏎ to Restart \n\n🔘 Close window to Exit\n\n").strip().lower()
            except EOFError:
                print("\n👋 Window closed. Exiting tool...")
                exit()
            if choice == "c":
                print("\n ✌ Allowed duplicates...")
            elif choice == "":
                restart_msg()
                continue
            else:
                restart_msg()
                continue

        print("\n🔘 Set validity from the pop-up calendar...")
        break

    # Convert the voucher column to integer
    df["Voucher"] = pd.to_numeric(df["Voucher"], errors="coerce").astype("Int64")
    
    # Sort by Voucher
    df = df.sort_values(by="Voucher")

    # Fix contact numbers
    df["Contact"] = df["Contact"].apply(lambda x: x if len(x) != 10 else "0" + x)

    # Get validity dates
    start_date, end_date = get_validity_dates()

    # Build segments
    segments = build_segments(df, start_date, end_date)
    final_text = f"Need to send notification for the coupon list below:\n\n{'\n'.join(segments)}"

    # Export as Notepad text file
    output_folder = os.path.join(os.path.expanduser("~"), "Desktop")
    file_name = f"Coupons_notification_{start_date.replace(' ', '_')}_to_{end_date.replace(' ', '_')}.txt"
    output_path = os.path.join(output_folder, file_name)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(final_text)

    # Successfull file generation reminder
    print("\n:::::: ✨  Notepad text file generated  ╰┈➤  📁", output_path)
    time.sleep(2)

    while True:
        try:
            choice = input("\n\n🔘 Press ENTER ⏎ to Restart \n\n🔘 Close window to Exit\n\n").strip()
        except EOFError:
            print("\n👋 Window closed. Exiting tool...")
            exit()
        if choice == "":
            os.system("cls" if os.name == "nt" else "clear")
            restart_msg()
            return main()
        else:
            os.system("cls" if os.name == "nt" else "clear")
            restart_msg()
            return main()

if __name__ == "__main__":
    main()
