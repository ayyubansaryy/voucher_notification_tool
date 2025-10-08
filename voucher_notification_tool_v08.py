    # Compilation command: pyinstaller --onefile "F:\__Practice\Python\voucher_notification_tool\voucher_notification_tool_v08.py"

import os
import time
from io import StringIO
import datetime
import tkinter as tk
import pandas as pd
from tkcalendar import Calendar
from reportlab.lib.pagesizes import A4
import re


def boot_msg():
    print("\n══════════════════════════════════════════════════════\nVoucher Notification Text Tool\n══════════════════════════════════════════════════════\nv0.7 (developed by: FEL-89242)\nDefault output location: Desktop\n══════════════════════════════════════════════════════\n")

def restart_msg():
    os.system("cls" if os.name == "nt" else "clear")
    print("\n🟢 Successfully Restarted\n\nPaste data below (with headers) & press 'Enter' twice:\n══════════════════════════════════════════════════════\n")

def closing():
    print("\n👋 Closing in 5 seconds...")
    time.sleep(3)

#### Format date to string
def format_date(date_input: str) -> str:
    try:
        return datetime.datetime.strptime(date_input, "%d/%m/%Y").strftime("%d %B")
    except ValueError:
        return date_input

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


#### Build notification text segments
def build_segments(df: pd.DataFrame, start: str, end: str) -> list[str]:
    segments = []

    def format_order_contact(line: str) -> str:
        match = re.match(r"(\w+)\s*(\d+)", line)
        if match:
            order, contact = match.groups()
            return f"{order}\u00A0\u00A0\u00A0{contact}"
        return line\
    
    # Split date into two segmenst (e.g. 20 january --> ["20", "January"])
    start_a = start.split()[0].lstrip("0")
    start_b = start.split()[1].lstrip("0")
    end_a = end.split()[0].lstrip("0")
    end_b = end.split()[1].lstrip("0")
    # prepare start day extension
    if start_a in ("1", "31", "21"): start_date = f"{start_a}st {start_b}"
    elif start_a in ("2", "22"): start_date = f"{start_a}nd {start_b}"
    elif start_a in ("3", "23"): start_date = f"{start_a}rd {start_b}"
    else: start_date = f"{start_a}th {start_b}"
    # prepare end day extension
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
            f"{serial}. {code_str}\n",
            *order_contact_lines,
            f"\nUse coupon {code_str} to get {int(amount)} taka off",
            f"Minimum order: {mov} taka",
            f"Validity: {start_date} to {end_date}",
            "Not applicable for Flat discount-providing restaurants\n",
        ]
        segments.append("\n".join(lines))
    return segments


#### Input Data Function
def read_input_data() -> pd.DataFrame:
    while True:
        lines = []
        while True:
            try:
                line = input()
                if line.strip() == "":
                    break
                lines.append(line)
            except EOFError:
                break

        raw = "\n".join(lines).strip()
        if not raw:
            print("🛑 No data found! The user did not provide any data!\n")
            
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

        # Detect header
        first_line = raw.splitlines()[0].strip().lower()
        has_header = any(keyword in first_line for keyword in ["ticket id", "order no", "contact", "voucher"])

        if has_header:
            df = pd.read_csv(StringIO(raw), sep="\t", dtype=str)

        else:
            print("🛑 ERROR! Headers not found!\n")
            try:
                choice = input("🔘 Press ENTER ⏎ to Restart \n\n🔘 Close window to Exit\n\n").strip()
            except EOFError:
                print("\nWindow closed. Exiting tool...")
                exit()
            if choice == "":
                restart_msg()
                continue
            else:
                restart_msg()
                continue

        # Exclude Voucher Given == yes or withdrawn
        if "Voucher Given" in df.columns:
            df = df[~df["Voucher Given"].astype(str).str.strip().str.lower().isin(["yes", "withdrawn"])]

        # Keep only relevant columns
        df = df[[col for col in ["Order No", "Contact", "Voucher"] if col in df.columns]]
        return df

#### Main Program
boot_msg()

choice = input("🟢 Set Voucher Session (M for Morning or E for Evening) : ").strip().lower()
if choice == "m" :
    user_session = "Morning"
    print(f"\n:::::: Voucher Session is set to '{user_session}'\n")

elif choice == "e" :
    user_session = "Evening"
    print(f"\n:::::: Voucher Session is set to '{user_session}'\n")

else: 
    user_session = "Default (Evening)"
    print(f"\n:::::: Voucher Session is set to '{user_session}'\n")

print("Paste data below (with headers) & press 'Enter' twice:\n══════════════════════════════════════════════════════\n")

def main():
    while True:
        df = read_input_data()
        if df is None:
            
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
                choice = input("\n🔘 Type '𝐂' press ENTER to continue \n\n🔘 Press ENTER ⏎ to Restart \n\n🔘 Close window to Exit\n\n").strip().lower()
            except EOFError:
                print("\nWindow closed. Exiting tool...")
                exit()
            if choice == "c":
                print("\n:::::: Allowed duplicates...")
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
    session_info = f"Voucher Session: {user_session}, Date: {start_date}\n\n"
    final_text = f"{session_info}\n\nNeed to send notification for the coupon list below:\n\n{'\n'.join(segments)}"

    # Export as Notepad text file
    output_folder = os.path.join(os.path.expanduser("~"), "Desktop")
    file_name = f"{user_session} {start_date} to {end_date}.txt"
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