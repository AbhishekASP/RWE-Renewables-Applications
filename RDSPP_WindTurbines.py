
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog


# =========================================================
# FILE PICKER (VS CODE COMPATIBLE)
# =========================================================
def pick_file():
    root = tk.Tk()
    root.withdraw()
    root.update()

    file_path = filedialog.askopenfilename(
        parent=root,
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    root.destroy()
    return file_path


# =========================================================
# YAW DRIVE COUNT POPUP
# =========================================================
def ask_yaw_drive_count():
    root = tk.Tk()
    root.withdraw()
    root.update()

    while True:
        val = simpledialog.askstring(
            "Yaw Drive Count",
            "Enter number of Yaw Drives (e.g., 3):",
            parent=root
        )

        if val is None:
            root.destroy()
            return None

        if val.isdigit() and int(val) >= 2:
            root.destroy()
            return int(val)

        messagebox.showerror("Invalid Input", "Please enter a number ≥ 2", parent=root)


# =========================================================
# CLEAN TEMPLATE / UNWANTED ROWS
# =========================================================
def clean_template_rows(df):

    remove_patterns = [
        "like mda", "like mdl", "same structure", "template",
        "mz0n0", "yaw drive n"
    ]
    pattern = "|".join(remove_patterns)

    df = df[
        ~df["Code"].str.contains(pattern, case=False, na=False) &
        ~df["Description (full)"].str.contains(pattern, case=False, na=False)
    ]

    # Remove stray RB2/3 root nodes from input:
    df = df[~df["Code"].isin(["MDA12", "MDA13"])]

    df = df.drop_duplicates(subset=["Code", "Description (full)"])

    return df


# =========================================================
# INSERT DATAFRAME AFTER INDEX
# =========================================================
def insert_after_index(data, new_data, index):
    return pd.concat(
        [data.iloc[:index + 1], new_data, data.iloc[index + 1:]],
        ignore_index=True
    )


# =========================================================
# ROTOR BLADE GENERATION
# =========================================================
def process_rotor_blades(data):

    rb1 = data[
        data["Code"].str.startswith("MDA11") &
        ~data["Description (full)"].str.contains("like", case=False, na=False)
    ]

    if rb1.empty:
        return data

    rb1_index = data[data["Code"].str.startswith("MDA11")].index.max()

    rb2 = rb1.copy()
    rb3 = rb1.copy()

    rb2["Code"] = rb2["Code"].str.replace("MDA11", "MDA12")
    rb3["Code"] = rb3["Code"].str.replace("MDA11", "MDA13")

    rb2["Description (full)"] = rb2["Description (full)"].str.replace("Rotor Blade 1", "Rotor Blade 2")
    rb3["Description (full)"] = rb3["Description (full)"].str.replace("Rotor Blade 1", "Rotor Blade 3")

    data = insert_after_index(data, rb2, rb1_index)
    rb2_index = data[data["Code"].str.startswith("MDA12")].index.max()
    data = insert_after_index(data, rb3, rb2_index)

    return data


# =========================================================
# FIND EXACT BLOCK OF YAW DRIVE 1
# =========================================================
def find_yaw_drive_1_block(data):

    start_idx = data[data["Code"] == "MDL10.MZ010"].index
    if start_idx.empty:
        return None, None

    start = start_idx[0]
    end = start

    for i in range(start + 1, len(data)):
        code = str(data.at[i, "Code"])

        # STOP when next yaw drive 2 appears
        if code.startswith("MDL10.MZ020"):
            break

        # STOP when other systems start
        if code.startswith(("MDL20", "MDX", "MDY", "MKA")):
            break

        # Otherwise include as child of YD1
        end = i

    return start, end


# =========================================================
# GENERATE YAW DRIVES 2..N
# =========================================================
def generate_yaw_drives(data, count):

    yaw1 = data[data["Code"].str.startswith("MDL10.MZ010")]

    if yaw1.empty:
        messagebox.showerror("Error", "MDL10.MZ010 (Yaw Drive 1) not found!")
        return pd.DataFrame()

    new_rows = []

    for i in range(2, count + 1):
        temp = yaw1.copy()

        new_suffix = f"MZ0{10*i:02d}"
        temp["Code"] = temp["Code"].str.replace("MZ010", new_suffix)

        temp["Description (full)"] = temp["Description (full)"].replace(
            {
                "Yaw Drive 1": f"Yaw Drive {i}",
                "Yaw Motor 1": f"Yaw Motor {i}",
                "Yaw Gear 1": f"Yaw Gear {i}",
            },
            regex=True
        )

        new_rows.append(temp)

    return pd.concat(new_rows, ignore_index=True)


# =========================================================
# INSERT YAW DRIVES 2..N AFTER YAW DRIVE 1 BLOCK
# =========================================================
def insert_yaw_drives(data, yaw_new):

    if yaw_new.empty:
        return data

    start, end = find_yaw_drive_1_block(data)
    if start is None:
        return data

    insert_pos = end + 1

    combined = pd.concat(
        [data.iloc[:insert_pos], yaw_new, data.iloc[insert_pos:]],
        ignore_index=True
    )

    return combined

# =========================================================
# REMOVE ANY STRAY YAW DRIVE 2..N APPEARING AFTER LAST DRIVE
# =========================================================
def remove_stray_yaw_drives(data, yaw_count):
    """
    Removes any accidental extra yaw drive blocks appearing AFTER
    the final expected yaw drive (MZ0xx end block).
    """

    # Expected final code prefix (e.g., MZ030 for yaw_count = 3)
    last_code = f"MZ0{10 * yaw_count:02d}"

    # Find last correct yaw drive row
    last_pos_list = data.index[data["Code"].str.startswith(f"MDL10.{last_code}")].tolist()

    if not last_pos_list:
        return data  # nothing to clean

    last_valid_pos = last_pos_list[-1]

    # Remove any MDL10.MZ0xx that occur AFTER last_valid_pos (stray duplicates)
    mask = data.index > last_valid_pos
    stray = data[mask & data["Code"].str.match(r"MDL10\.MZ0[0-9]+0", na=False)]

    if not stray.empty:
        data = data.drop(stray.index)

    return data.reset_index(drop=True)


# =========================================================
# FILTER AFTER FULL TREE READY
# =========================================================
def filter_and_preserve_hierarchy(data):

    data["Level"] = data["Level"].astype(str)
    data["Marked"] = data["Comment"].apply(lambda x: "x" in str(x).lower())
    data["To Keep"] = False

    for idx, row in data.iterrows():
        if row["Marked"]:
            data.at[idx, "To Keep"] = True

            level = row["Level"]
            current = idx

            while level.startswith("F") and level[1:].isdigit() and int(level[1:]) > 0:
                current -= 1
                parent_level = data.at[current, "Level"]

                if parent_level.startswith("F") and parent_level[1:].isdigit() and int(parent_level[1:]) < int(level[1:]):
                    data.at[current, "To Keep"] = True
                    level = parent_level

    return data[data["To Keep"]].drop(columns=["Marked", "To Keep"])


# =========================================================
# MAIN FLOW
# =========================================================
def main():

    file = pick_file()
    if not file:
        print("❌ No file selected.")
        return

    yaw_count = ask_yaw_drive_count()
    if yaw_count is None:
        print("❌ No yaw count entered.")
        return

    print("⏳ Processing...")

    data = pd.read_excel(file)
    data.columns = data.columns.str.strip()

    data = clean_template_rows(data)
    data = process_rotor_blades(data)

    # FULL TREE BEFORE FILTERING
    yaw_new = generate_yaw_drives(data, yaw_count)
    full_tree = insert_yaw_drives(data, yaw_new)

    # ✅ NEW FIX — REMOVE STRAY YAW DRIVE 2..N
    full_tree = remove_stray_yaw_drives(full_tree, yaw_count)

    full_tree.to_excel("Output1.xlsx", index=False)
    print("✔ Output1.xlsx generated (full tree).")
    # NOW FILTER
    final = filter_and_preserve_hierarchy(full_tree)
    final.to_excel("FinalOutput.xlsx", index=False)
    print("🎉 FinalOutput.xlsx generated successfully!")



if __name__ == "__main__":
    main()
