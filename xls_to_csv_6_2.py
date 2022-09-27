import multiprocessing
from multiprocessing import freeze_support
import pandas as pd
import tkinter as tk
import tkinter.filedialog
import tkinter.messagebox
import sys
import io
import os
import json
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
import threading
import warnings


def watchdog_monitor():

    def on_created(event):
        print(type(event))
        print(event.src_path)

        if event.src_path[-4:] == ".xls" or event.src_path[-5:] == ".xlsx":
            print("true")
            time.sleep(1)
            confirm(event.src_path)

    json_f = json.load(open("tmp/config.json"))
    event_handler = FileSystemEventHandler()

    event_handler.on_created = on_created

    path = json_f['defaultPath']

    observer = Observer()
    observer.schedule(event_handler, path, recursive=True)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()

    observer.join()


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
    global x
    directory = tk.filedialog.askdirectory()
    json_file["defaultPath"] = directory

    with open("tmp/config.json", "w") as f:
        json.dump(json_file, f, indent=4)

    x.terminate()

    x = multiprocessing.Process(target=watchdog_monitor, args=())
    x.start()


def confirm(filename):
    filename = f"{filename}"

    # print(filename)

    try:
        df = pd.read_excel(filename, engine="openpyxl")

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
            # tk.messagebox.showerror(f"Error: {e} not found", f"{e}")
            print(e)

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

            df_select["Quantity"] = df_select["Quantity"].astype(int)

            # print(df_select)

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
                        df_select.to_csv(filename[:-3] + f"{current_value}." + "csv", index=False, encoding='utf-8')
                    except OSError as e:
                        # tk.messagebox.showerror("Error: OSError", f"{e}")
                        print(e)
                        continue
                else:
                    # tk.messagebox.showerror("Error: OSError", f"{e}")
                    continue

    except Exception as e:
        print(e)


def on_close():
    # custom close options, here's one example:

    root.destroy()
    x.terminate()
    sys.exit()


if __name__ == "__main__":
    freeze_support()

    try:
        os.makedirs("tmp/")
    except FileExistsError:
        # print("dir exist")
        pass


    root = tk.Tk()
    root.title("Convert xls to csv auto")
    root.geometry("500x160+350+250")

    label1 = None
    df = None
    df_group = None
    df_group_list = [""]

    root.protocol("WM_DELETE_WINDOW", on_close)

    startup_check("tmp/config.json")

    json_file = json.load(open('tmp/config.json'))

    bt1 = tk.Button(root, text="Select default directory", command=select_default).pack()

    startup_check("tmp/config.json")

    x = multiprocessing.Process(target=watchdog_monitor, args=())
    x.start()

    root.mainloop()