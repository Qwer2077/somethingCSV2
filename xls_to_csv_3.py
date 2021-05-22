import pandas as pd
import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
import io
import os
import json

root = tk.Tk()
root.title("Convert xls to csv")
root.geometry("500x160+350+250")

filename = ""
label1 = None
df = None
df_group = None
df_group_list = [""]

try:
    os.makedirs("tmp/")
except FileExistsError:
    # print("dir exist")
    pass


def startup_check(file_path):
    if os.path.isfile(file_path) and os.access(file_path, os.R_OK):
        # checks if file exists
        # print("File exists and is readable")
        pass
    else:
        # print("Either file is missing or is not readable, creating file...")
        with open(file_path, 'w+') as fp:
            config = {
                "defaultPath": "/"
            }

            json.dump(config, fp, sort_keys=True, indent=4)


def select_default():
    directory = tk.filedialog.askdirectory()
    json_file["defaultPath"] = directory

    with open("tmp/config.json", "w") as f:
        json.dump(json_file, f, indent=4)


def openfile():
    global filename, label1, df, df_group, df_group_list

    if label1 is not None:
        label1.destroy()

    filename = tk.filedialog.askopenfilename(initialdir=json_file["defaultPath"], title="Select a xls file",
                                             filetypes=(("xls files", "*.xls"), ("all files", "*.*")))

    label1 = tk.Label(root, text=filename)

    if filename == "":
        return

    df = pd.read_excel(filename)
    df.rename(columns={
        "Unnamed: 2": "Glass Type",
        "Unnamed: 1": "Job Type"
    }, inplace=True)

    df[["Height", "Width"]] = df["Size"].str.split("x", expand=True)

    cols = df.columns.tolist()

    # df = df[cols]
    # print(cols)

    try:
        df = df[["Quantity", "Height", "Width", "Marks / Code", "Glass Type", "Job Type"]]
    except KeyError as e:
        tk.messagebox.showerror(f"Error: {e} not found", f"{e}")
        return

    df["Quantity"] = df["Quantity"].replace("x ", "", regex=True)

    df["Job Type"] = df["Job Type"].ffill()

    # df_group = df.groupby("Glass Type")

    df_group = df.groupby("Job Type")

    # print(df_group.groups.keys())

    df_group_list = list(df_group.groups.keys())

    label1.pack()

    # for name, names in df_group:
    #     print(name)
    #     print(names)
    #     print()

    value_inside.set("Select Glass Type")

    option_button1['menu'].delete(0, "end")
    for item in df_group_list:
        option_button1["menu"].add_command(label=item, command=tk._setit(value_inside, item))


def confirm():
    global value_inside, filename

    try:
        if df_group is None:
            raise FileNotFoundError

        df_select = df_group.get_group(value_inside.get())

        df_select = df_select[df_select["Glass Type"].notna()]

    except KeyError:
        tk.messagebox.showerror("Error: Glass type not selected", "Please select a glass type")
        return
    except FileNotFoundError:
        tk.messagebox.showerror("Error: File not selected", "Please select a file")
        return

    try:
        df_select.to_csv(filename[:-3] + f"{value_inside.get()}." + "csv", index=False)
    except OSError as e:

        error_list = list('\\/:*?"<>|')

        if any(x in error_list for x in value_inside.get()):
            current_value = value_inside.get()
            current_value = current_value.replace('\\', "_")
            current_value = current_value.replace("/", "_")
            current_value = current_value.replace(":", "_")
            current_value = current_value.replace("*", "_")
            current_value = current_value.replace("?", "_")
            current_value = current_value.replace("<", "_")
            current_value = current_value.replace(">", "_")
            current_value = current_value.replace("|", "_")

            try:
                df_select.to_csv(filename[:-3] + f"{current_value}." + "csv", index=False)
            except OSError:
                tk.messagebox.showerror("Error: OSError", f"{e}")
                return
        else:
            tk.messagebox.showerror("Error: OSError", f"{e}")
            return


startup_check("tmp/config.json")

json_file = json.load(open('tmp/config.json'))

value_inside = tk.StringVar(root)
value_inside.set("Select an Glass Type")

bt1 = tk.Button(root, text="Select default directory", command=select_default).pack()
bt2 = tk.Button(root, text="Open File", command=openfile).pack()
option_button1 = tk.OptionMenu(root, value_inside, *df_group_list)
option_button1.pack()

# print(df_group_list)

bt3 = tk.Button(root, text="Export", command=confirm).pack()

root.mainloop()
