import pyautogui
import keyboard
import pydirectinput
import time
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


def confirm():
    # keyboard.wait("esc")
    #
    # pyautogui.click(x=592, y=37)
    # pyautogui.hotkey("ctrl", "c")
    # pyautogui.click(x=316, y=856, button="right")
    # pyautogui.move(10, -340)
    # time.sleep(0.40)
    # pyautogui.click(x=559, y=584)
    # time.sleep(0.30)
    # pyautogui.hotkey("ctrl", "v")
    # keyboard.press("enter")

    keyboard.wait("esc")
    pydirectinput.PAUSE = False

    pyautogui.click(x=592, y=37)

    # pyautogui.hotkey("ctrl", "c")
    # keyboard.on_press_key("ctrl")
    # keyboard.press("c")
    pydirectinput.keyDown("ctrl")
    pydirectinput.press("c")
    pydirectinput.keyUp("ctrl")

    pyautogui.click(x=316, y=856, button="right")
    pyautogui.moveTo(378, 632)
    time.sleep(0.40)
    pyautogui.click(x=534, y=682)
    time.sleep(1.5)

    # pyautogui.hotkey("ctrl", "v")
    pydirectinput.keyDown("ctrl")
    pydirectinput.press("v")
    pydirectinput.keyUp("ctrl")

    keyboard.press("enter")

    time.sleep(0.5)

    filename = f"{json_file['defaultPath']}/{root.clipboard_get()}.xls"

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

    for glass_type in df_group_list:
        df_select = df_group.get_group(glass_type)

        df_select = df_select[df_select["Glass Type"].notna()]

        try:
            df_select.to_csv(filename[:-3] + f"{glass_type}." + "csv", index=False)
        except OSError as e:

            error_list = list('\\/:*?"<>|')

            if any(x in error_list for x in glass_type):
                current_value = glass_type
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

bt1 = tk.Button(root, text="Select default directory", command=select_default).pack()
bt2 = tk.Button(root, text="开始转换", command=confirm).pack()


root.mainloop()

# keyboard.wait("esc")

# print(pyautogui.position())


# pyautogui.click(x=592, y=37)
# pyautogui.hotkey("ctrl", "c")
# pyautogui.click(x=319, y=869, button="right")
# pyautogui.move(10, -340)
# time.sleep(0.30)
# pyautogui.click(x=559, y=584)
# time.sleep(0.30)
# pyautogui.hotkey("ctrl", "v")
# keyboard.press("enter")
