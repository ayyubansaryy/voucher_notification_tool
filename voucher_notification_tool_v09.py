    # Compilation command: pyinstaller --onefile --distpath . "F:\__Practice\Python\voucher_notification_tool\voucher_notification_tool_v09.py"

import os
import time
from io import StringIO
import datetime
import tkinter as tk
import pandas as pd
from tkcalendar import Calendar
from reportlab.lib.pagesizes import A4
import re
from tabulate import tabulate

version = "0.9"

# --------------------------- Boot Messages --------------------------- #

def boot_msg():
    print("\n" + "═" * 57)
    print("🟢  VOUCHER NOTIFICATION TEXT TOOL 🟢".center(57))
    print("-" * 57)
    print(f"Version: {version}".center(57))
    print("FEL-89242".center(58))
    print("-" * 57)
    print("╰┈➤  Default output location: Desktop".center(57))
    print("═" * 57 + "\n")

def restart_msg():
    os.system("cls" if os.name == "nt" else "clear")
    print("\n🟢 Restarted\n\nPaste data below (with headers) & press 'Enter' twice:" + "\n" + "-" * 54 + "\n")

def style(msg):
    print(("+" * (len(msg)+10)) + f"\n{msg}\n" + ("+" * (len(msg)+10)) + "\n")

def closing():
    print("\nClosing in 5 seconds...")
    time.sleep(3)


# --------------------------- Helpers --------------------------- #

def format_date(date_input: str) -> str:
    try:
        return datetime.datetime.strptime(date_input, "%d/%m/%Y").strftime("%d %B")
    except ValueError:
        return date_input

def restart_or_exit():
    try:
        choice = input("\n🔘 press ENTER to Restart \n\n🔘 close window to EXIT\n\n").strip()
    except EOFError:
        print("\nWindow closed. Exiting tool...")
        exit()
    restart_msg()
    return choice == ""


# --------------------------- Calendar --------------------------- #
def pick_date(title="Select a date") -> str:
    root = tk.Tk()
    root.title(title)
    root.configure(bg="#F8F9FA")

    tk.Label(root, text=title, font=("Segoe UI", 16, "bold"), bg="#F8F9FA").pack(pady=10)

    cal = Calendar(root, selectmode="day", date_pattern="dd/mm/yyyy", font=("Segoe UI", 14))
    cal.pack(pady=10)

    w, h = 500, 400
    x = (root.winfo_screenwidth() - w) // 2
    y = (root.winfo_screenheight() - h) // 2
    root.geometry(f"{w}x{h}+{x}+{y}")

    selected_date = tk.StringVar()

    def confirm():
        selected_date.set(cal.get_date())
        root.destroy()

    tk.Button(
        root,
        text="Confirm ✅",
        font=("Segoe UI", 12, "bold"),
        bg="#0067C7",
        fg="white",
        relief="flat",
        padx=20,
        pady=8,
        command=confirm
    ).pack(pady=15)

    root.lift()
    root.attributes("-topmost", True)
    root.after(100, lambda: root.attributes("-topmost", False))
    root.mainloop()
    return selected_date.get()

def get_validity_dates():
    start = format_date(pick_date("Validity Starts:"))
    end = format_date(pick_date("Validity Ends:"))
    return start, end

# --------------------------- Build Text Segments --------------------------- #
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

# --------------------------- Data Reader --------------------------- #
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
            msg = "🔴 No data found! The user did not provide any data!"
            style(msg)

            if restart_or_exit():
                continue
            else:
                continue

        first_line = raw.splitlines()[0].strip().lower()
        has_header = any(k in first_line for k in ["date", "ticket no", "order no", "contact", "voucher", "voucher given"])

        if not has_header:
            msg = "🔴 ERROR! Headers not found!"
            style(msg)

            if restart_or_exit():
                continue
            else:
                continue

        df = pd.read_csv(StringIO(raw), sep="\t", dtype=str)

        if "Voucher Given" in df.columns:
            df = df[~df["Voucher Given"].astype(str).str.strip().str.lower().isin(["yes", "withdrawn"])]
            
            if df is None or df.empty:
                msg = "🔴 All vouchers have already been given!"
                style(msg)

        df = df[[col for col in ["Order No", "Contact", "Voucher"] if col in df.columns]]
        return df

# --------------------------- Table Preview --------------------------- #
def preview_table(df: pd.DataFrame) -> bool:

    print("\n🔹 Preview of Pasted Data:")

    # Show table
    try:
        from tabulate import tabulate
        print(tabulate(df, headers="keys", tablefmt="grid", showindex=False))
    except ImportError:
        with pd.option_context("display.max_rows", None, "display.max_columns", None):
            print(df.to_string(index=False))

    # Show summary info
    total_rows = len(df)
    print(f"\n🔹 Total Valid Entries: {total_rows}\n")

    # Show voucher counts, sorted numerically
    if "Voucher" in df.columns:
        df["Voucher"] = pd.to_numeric(df["Voucher"], errors="coerce")
        voucher_counts = df["Voucher"].value_counts().sort_index(ascending=True)

        print("🔹 Entries per Voucher:\n")
        for voucher, count in voucher_counts.items():
            print(f"   {int(voucher)}  ┈┈┈>  {count}")

    # Confirmation loop
    while True:
        try:
            choice = input(
                "\n🔘 press ENTER to continue\n\n"
                "🔘 type 'R' and press ENTER to restart tool\n\n "
            ).strip().lower()

            if choice == "r":
                restart_msg()
                return False
            elif choice == "":
                return True
            else:
                msg = "🔴 Invalid input — please press ENTER or type 'R'"
                style(msg)

        except (KeyboardInterrupt, EOFError):
            print("\n🔴 Execution interrupted.\n\n🟢 Restarting tool...\n")
            restart_msg()
            return False

# --------------------------- Main Program --------------------------- #
def main():
    boot_msg()

    choice = input("🔘 Set Voucher Session (M for Morning or E for Evening):\n\n ").strip().lower()
    if choice == "m":
        user_session = "Morning"
    elif choice == "e":
        user_session = "Evening"
    else:
        user_session = "Evening (Default)"

    print(f"\n:::::: Voucher session has been set to '{user_session}'\n")
    print("Paste data below (with headers) & press 'Enter' twice:" + "\n" + "-" * 57 + "\n")

    while True:
        df = read_input_data()
        if df is None or df.empty:
            if restart_or_exit():
                continue
            else:
                continue

        # Show pasted data as table preview
        if not preview_table(df):
            continue

        # Check for rows with missing Voucher
        missing_voucher_rows = df[df["Voucher"].isna()]
        if not missing_voucher_rows.empty:
            print("\n🔴 ERROR! Following rows have missing Voucher Amount:\n")
            for _, row in missing_voucher_rows.iterrows():
                print(f"{row['Order No']}  0{row['Contact']}  {row['Voucher']}")

            if restart_or_exit():
                continue
            else:
                continue

        # Check for rows with missing Order Number
        missing_order_rows = df[df["Order No"].isna() | (df["Order No"].astype(str).str.strip() == "")]
        if not missing_order_rows.empty:
            print("\n🔴 ERROR! Following rows have Missing Order Number:\n")

            for _, row in missing_order_rows.iterrows():
                print(f"Contact: 0{row['Contact']}  Voucher: {row['Voucher']}")

            if restart_or_exit():
                continue
            else:
                continue

        # Check for rows with duplicate Contact
        duplicate_contacts = df[df.duplicated(subset=["Contact"], keep=False)]
        if not duplicate_contacts.empty:
            duplicate_contacts = duplicate_contacts.sort_values(by="Contact")
            print("\n🟠 Warning! Duplicate contacts found:\n")
            for _, row in duplicate_contacts.iterrows():
                print(f"{row['Order No']} 0{row['Contact']} {int(row['Voucher'])}")

            try:
                choice = input(
                    "\n🔘 Press ENTER to continue \n\n🔘 Type 'R' and press ENTER to Restart \n\n🔘 Close window to Exit\n\n"
                ).strip().lower()
            except EOFError:
                print("\nWindow closed. Exiting tool...")
                exit()

            if choice == "r":
                restart_msg()
                continue
            else:
                print(":::::: Duplicates allowed\n")

        print("\n🔘 Set validity from the pop-up calendar...")

        # Fix contact number (add "0" before each one)
        df["Contact"] = df["Contact"].apply(lambda x: x if len(x) != 10 else "0" + x)

        # Make vouchers integers
        df["Voucher"] = pd.to_numeric(df["Voucher"], errors="coerce").fillna(0).astype(int)

        # Sort rows by the voucher amounts
        df = df.sort_values(by="Voucher")

        # Get validity
        start_date, end_date = get_validity_dates()

        # Build segments
        segments = build_segments(df, start_date, end_date)

        # setup output file
        output_folder = os.path.join(os.path.expanduser("~"), "Desktop")
        file_name = f"{user_session}_{start_date.replace(" ", "_")}_to_{end_date.replace(" ", "_")}.txt"
        output_path = os.path.join(output_folder, file_name)

        with open(output_path, "w", encoding="utf-8") as f:
            f.write(f"\nNeed to send notification for the coupon list below:\n\n{'\n\n'.join(segments)}")

        # Successful file generation message
        print()
        msg = f"\n:::::: ✨ Notepad text file generated  ╰┈➤  📁 {output_path}\n"
        style(msg)

        # Auto open the file in Notepad on Windows
        if os.name == "nt":
            os.system(f'notepad "{output_path}"')

        time.sleep(2)

        if restart_or_exit():
            continue
        else:
            continue


if __name__ == "__main__":
    main()