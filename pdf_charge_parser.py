import pandas as pd
import fitz
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import StringVar
import threading
import time
from PIL import Image, ImageTk
import os


# Clean Phone and Username
def extract_phone_user(df, col):
    df[['Phone', 'Username']] = df[col].str.extract(r'(\d+)\|\|\|(.+)', expand=True)
    df['Phone'] = df['Phone'].str.replace(r'\D', '', regex=True)
    return df

# PDF Reader
def pdf_reader(pdf_path, roaming_df, country_df, update_progress):
    country_map = dict(zip(country_df['Acronym'].str.strip().str.upper(), country_df['Country Name']))
    doc = fitz.open(pdf_path)
    results = []
    total = len(roaming_df)
    for idx, (phone, name, amt) in enumerate(roaming_df[['Phone', 'Username', 'Amount']].values):
        formatted_phone = f"{phone[:3]}-{phone[3:6]}-{phone[6:]}"
        for page in doc:
            text = page.get_text()
            if formatted_phone in text:
                matches = re.findall(r"Mobile Browser:\s*([A-Z]{3})", text, re.IGNORECASE)
                for code in matches:
                    if code.strip().upper() in country_map:
                        results.append([formatted_phone, name, country_map[code.strip().upper()], amt])
        update_progress("PDF", round((idx + 1) / total * 100, 2))

    df_results = pd.DataFrame(results, columns=["Phone Number", "Username", "Charge Detail", "Charge Amount"])
    df_results = df_results.drop_duplicates()
    df_results = (
        df_results.groupby(['Phone Number', 'Username', 'Charge Amount'])
        .agg({'Charge Detail': lambda x: 'Roaming to ' + ' & '.join(sorted(set(x)))})
        .reset_index()
    )
    return df_results

# LD Reader
def ld_reader(pdf_path, ld_df):
    doc = fitz.open(pdf_path)
    phone_dest_map = {}
    for page in doc:
        lines = [line.strip() for line in page.get_text().split('\n') if line.strip()]
        phone = None
        for i in range(len(lines) - 1):
            if lines[i] == "Mobile":
                phone = lines[i + 1].replace('-', '')
                break
        if phone not in ld_df['Phone'].values:
            continue

        prefix = "Roaming LD to" if any("Roamer" in l for l in lines) else "LD to"
        try:
            idx = lines.index("ITEMIZED LONG DISTANCE CALLS")
        except:
            continue
        section = lines[idx+1:]
        data_start = next((i for i, x in enumerate(section) if x.isdigit()), None)
        if data_start is None:
            continue

        blocks, block = [], []
        for line in section[data_start:]:
            if line.lower().startswith("total"): break
            if line.strip().isdigit() and block:
                blocks.append(block)
                block = [line]
            else:
                block.append(line)
        if block: blocks.append(block)

        for block in blocks:
            try:
                nums = [float(x) for x in block if x.replace('.', '', 1).isdigit()]
                if not nums or nums[-1] <= 0:
                    continue
                dest = next(x for x in block[::-1] if not x.replace('.', '', 1).isdigit())
                if phone not in phone_dest_map:
                    phone_dest_map[phone] = {"to": set(), "prefix": prefix}
                phone_dest_map[phone]["to"].add(dest)
            except:
                continue

    records = []
    for _, row in ld_df.iterrows():
        phone = row['Phone']
        name = row['Username']
        amount = row['Amount']
        fallback = row['Description']
        formatted = f"{phone[:3]}-{phone[3:6]}-{phone[6:]}"
        if phone in phone_dest_map:
            destinations = sorted(phone_dest_map[phone]["to"])
            prefix = phone_dest_map[phone]["prefix"]
            charge_detail = f"{prefix}: {' & '.join(destinations)}"
        else:
            charge_detail = fallback
        records.append([formatted, name, charge_detail, amount])
    return pd.DataFrame(records, columns=["Phone Number", "Username", "Charge Detail", "Charge Amount"])

# Roaming LD Reader
def roaming_ld_reader(pdf_path, aux_df):
    phone_map = dict(zip(aux_df['Phone'], aux_df['Username']))
    doc = fitz.open(pdf_path)
    records = []

    for page in doc:
        lines = [line.strip() for line in page.get_text().split('\n') if line.strip()]
        total_lines = [l for l in lines if l.lower().startswith("total")]
        total_amt = 0
        for l in total_lines:
            try:
                val = lines[lines.index(l) + 1].replace('$', '')
                total_amt = float(val)
                if total_amt > 0: break
            except: continue
        if total_amt <= 0 or "Roamer" not in lines or "ITEMIZED LONG DISTANCE CALLS" not in lines:
            continue

        phone = None
        for i in range(len(lines)-1):
            if lines[i] == "Mobile":
                phone = re.sub(r'\D', '', lines[i + 1])
                break
        if not phone: continue

        idx = lines.index("ITEMIZED LONG DISTANCE CALLS")
        section = lines[idx+1:]
        data_start = next((i for i, x in enumerate(section) if x.isdigit()), None)
        if data_start is None: continue

        blocks, block = [], []
        for line in section[data_start:]:
            if line.lower().startswith("total"): break
            if line.strip().isdigit() and block:
                blocks.append(block)
                block = [line]
            else:
                block.append(line)
        if block: blocks.append(block)

        for block in blocks:
            try:
                nums = [float(x) for x in block if x.replace('.', '', 1).isdigit()]
                if not nums or nums[-1] <= 0: continue
                charge = nums[-1]
                dest = next(x for x in block[::-1] if not x.replace('.', '', 1).isdigit())
                formatted = f"{phone[:3]}-{phone[3:6]}-{phone[6:]}"
                records.append([formatted, phone_map.get(phone, "Unknown"), f"Roaming LD to: {dest}", charge])
            except:
                continue
    return pd.DataFrame(records, columns=["Phone Number", "Username", "Charge Detail", "Charge Amount"])

def main_ui():
    app = tk.Tk()
    app.title("Expertel Telecom")
    app.geometry("700x600")

    # Load and display logo
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    logo_path = os.path.join(script_dir, "logo.png")
    logo = Image.open(logo_path)
    logo = logo.resize((200, 60), Image.LANCZOS)
    logo_img = ImageTk.PhotoImage(logo)

    logo_label = tk.Label(app, image=logo_img)
    logo_label.image = logo_img
    logo_label.pack(pady=(15, 5))

    # Header Text
    header = tk.Label(app, text="Additional Charges Report Generator", font=("Arial", 14, "bold"))
    header.pack(pady=(0, 15))

    paths = {"country": StringVar(), "roaming": StringVar(), "ld": StringVar(), "pdf": StringVar(), "output": StringVar()}
    progress_bars = {}
    percent_labels = {}

    
    # Wrapping everything inside a centered frame
    container = tk.Frame(app)
    container.place(relx=0.5, rely=0.5, anchor="center")

    def browse_file(var, types):
        path = filedialog.askopenfilename(filetypes=types)
        if path: var.set(path)

    def browse_output():
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[["Excel files", "*.xlsx"]])
        if path: paths["output"].set(path)

    def update_progress(stage, value):
        progress_bars[stage]["value"] = value
        percent_labels[stage].config(text=f"{int(value)}%")
        app.update_idletasks()

    def run_processing():
        def task():
            try:
                df_country = pd.read_csv(paths['country'].get())
                update_progress("Country", 100)
                time.sleep(0.5)

                df_roaming = extract_phone_user(pd.read_excel(paths['roaming'].get()), 'Phone')
                update_progress("Roaming", 100)
                time.sleep(0.5)

                df_ld = extract_phone_user(pd.read_excel(paths['ld'].get()), 'User Name')
                raw_ld = pd.read_excel(paths['ld'].get())
                df_ld['Description'] = raw_ld['Description']
                df_ld['Amount'] = raw_ld['Amount']
                update_progress("LD", 100)
                time.sleep(0.5)

                df1 = pdf_reader(paths['pdf'].get(), df_roaming, df_country, update_progress)
                df2 = ld_reader(paths['pdf'].get(), df_ld)
                update_progress("PDF", 90)
                df3 = roaming_ld_reader(paths['pdf'].get(), df_roaming)
                update_progress("PDF", 100)
                time.sleep(0.5)

                df_final = pd.concat([df1, df2, df3], ignore_index=True)
                df_final.to_excel(paths['output'].get(), index=False)
                update_progress("Final", 100)
                time.sleep(0.5)

                messagebox.showinfo("Done", f"Output saved at:\n{paths['output'].get()}")
            except Exception as e:
                messagebox.showerror("Error", str(e))

        threading.Thread(target=task).start()

    # Form Inputs
    row = 0
    for label, key, types in [
        ("1. Country List (.csv)", "country", [["CSV Files", "*.csv"]]),
        ("2. Roaming Excel (.xlsx)", "roaming", [["Excel Files", "*.xlsx"]]),
        ("3. Additional Charges Excel (.xlsx)", "ld", [["Excel Files", "*.xlsx"]]),
        ("4. Monthly PDF File (.pdf)", "pdf", [["PDF Files", "*.pdf"]])
    ]:
        tk.Label(container, text=label).grid(row=row, column=0, sticky='w', padx=10, pady=5)
        tk.Entry(container, textvariable=paths[key], width=50).grid(row=row, column=1)
        tk.Button(container, text="Browse", command=lambda v=paths[key], t=types: browse_file(v, t)).grid(row=row, column=2)
        row += 1

    tk.Label(container, text="5. Save Output As").grid(row=row, column=0, sticky='w', padx=10, pady=5)
    tk.Entry(container, textvariable=paths['output'], width=50).grid(row=row, column=1)
    tk.Button(container, text="Browse", command=browse_output).grid(row=row, column=2)
    row += 1

    # Progress Bars
    for stage in ["Country", "Roaming", "LD", "PDF", "Final"]:
        tk.Label(container, text=f"{stage} Progress:").grid(row=row, column=0, sticky='e', padx=10)
        pbar = ttk.Progressbar(container, orient="horizontal", length=300, mode="determinate")
        pbar.grid(row=row, column=1, pady=5)
        lbl = tk.Label(container, text="0%")
        lbl.grid(row=row, column=2)
        progress_bars[stage] = pbar
        percent_labels[stage] = lbl
        row += 1

    # Button
    tk.Button(container, text="Create Report", command=run_processing, bg='green', fg='white').grid(row=row, column=1, pady=20)

    app.mainloop()


# Run the app
if __name__ == "__main__":
    main_ui()
