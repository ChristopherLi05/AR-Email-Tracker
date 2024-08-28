import csv
import io
import json
import os
import sys
import data_parser
import tkinter as tk
from tkinter.filedialog import askopenfilename, asksaveasfilename
from datetime import date


class MainFrameIO(io.StringIO):
    def __init__(self, mainframe: 'MainFrame'):
        super().__init__()
        self.parent = mainframe

    def write(self, __s):
        io.StringIO.write(self, __s)
        self.parent.stdout(__s)


class MainFrame(tk.Frame):
    def __init__(self, parent):
        tk.Frame.__init__(self, parent)

        self.manager = data_parser.TrackerManager(add_dummy=True)

        self.lock_input = False
        self.lock_extract = False

        self.unknown_emails = {}

        # Hijacking print console
        self.orig_stdout = sys.stdout
        sys.stdout = MainFrameIO(self)

        tk.Label(text="Tracker File: ").grid(row=0, column=0, sticky="e")
        self.tracker_btn = tk.Button(parent, text="Pick File", command=self.load_tracker_cb)
        self.tracker_file = None
        self.tracker_btn.grid(row=0, column=1, sticky="w")

        # tk.Label(text="Tracker Entries: ", font=(font.nametofont('TkTextFont').actual(), 7)).grid(row=1, column=0, sticky="e")
        tk.Label(text="Tracker Entries: ").grid(row=10, column=0, sticky="e")
        self.tracker_count = tk.Label(text="0")
        self.tracker_count.grid(row=10, column=1, sticky="w")

        tk.Label(text="Email Mapping File: ").grid(row=2, column=0, sticky="e")
        self.email_mapping_btn = tk.Button(parent, text="Pick File", command=self.load_email_mapping_cb)
        self.email_mapping_file = None
        self.email_mapping_btn.grid(row=2, column=1, sticky="w")

        # tk.Label(text="Email Mapping Entries: ", font=(font.nametofont('TkTextFont').actual(), 7)).grid(row=3, column=0, sticky="e")
        tk.Label(text="Email Mapping Entries: ").grid(row=11, column=0, sticky="e")
        self.email_mapping_count = tk.Label(text="0")
        self.email_mapping_count.grid(row=11, column=1, sticky="w")

        tk.Label(text="Blacklist File: ").grid(row=4, column=0, sticky="e")
        self.blacklist_btn = tk.Button(parent, text="Pick File", command=self.load_blacklist_cb)
        self.blacklist_file = None
        self.blacklist_btn.grid(row=4, column=1, sticky="w")

        # tk.Label(text="Blacklist Entries: ", font=(font.nametofont('TkTextFont').actual(), 7)).grid(row=5, column=0, sticky="e")
        tk.Label(text="Blacklist Entries: ").grid(row=12, column=0, sticky="e")
        self.blacklist_count = tk.Label(text="0")
        self.blacklist_count.grid(row=12, column=1, sticky="w")

        self.load_files_btn = tk.Button(text="Load Files", width=20, command=self.load_files)
        self.load_files_btn.grid(row=6, column=1)
        tk.Button(text="Reset Data", width=20, command=self.reset_data_cb).grid(row=6, column=2)

        # self.info_label = tk.Label(text="Loaded Tracker Entries: 0 | Loaded Mapping Entries: 0 | Loaded Blacklist Entries: 0")
        # self.info_label.grid(row=7, column=0, columnspan=4)

        tk.Label(text="").grid(row=9, column=0)
        tk.Label(text="").grid(row=14, column=0)

        tk.Label(text="Email File: ").grid(row=15, column=0, sticky="e")
        self.email_btn = tk.Button(parent, text="Pick File", command=self.load_email_cb)
        self.email_file = None
        self.email_btn.grid(row=15, column=1, sticky="w")

        tk.Button(text="Run", width=20, command=self.run_cb).grid(row=16, column=1)
        self.ext_map_btn = tk.Button(text="Extract Mappings", width=20, command=self.ext_map_cb, state="disabled")
        self.ext_map_btn.grid(row=16, column=2)
        self.ext_tot_btn = tk.Button(text="Extract Total", width=20, command=self.ext_tot_cb, state="disabled")
        self.ext_tot_btn.grid(row=17, column=1)
        self.ext_week_btn = tk.Button(text="Extract Weekly", width=20, command=self.ext_week_cb, state="disabled")
        self.ext_week_btn.grid(row=17, column=2)

    def load_tracker_cb(self):
        if self.lock_input:
            return

        self.tracker_file = askopenfilename(filetypes=[("CSV Files", "*.csv")], initialdir=os.path.dirname(__file__), title="Pick tracker file")

        if self.tracker_file:
            self.tracker_btn.config(text=os.path.basename(self.tracker_file))
        else:
            self.tracker_btn.config(text="Pick File")

    def load_email_cb(self):
        self.email_file = askopenfilename(filetypes=[("PST Files", "*.pst")], initialdir=os.path.dirname(__file__), title="Pick email export file")

        if self.email_file:
            self.email_btn.config(text=os.path.basename(self.email_file))
        else:
            self.email_btn.config(text="Pick File")

    def load_email_mapping_cb(self):
        if self.lock_input:
            return

        self.email_mapping_file = askopenfilename(filetypes=[("JSON Files", "*.json")], initialdir=os.path.dirname(__file__), title="Pick email mapping file")

        if self.email_mapping_file:
            self.email_mapping_btn.config(text=os.path.basename(self.email_mapping_file))
        else:
            self.email_mapping_btn.config(text="Pick File")

    def load_blacklist_cb(self):
        if self.lock_input:
            return

        self.blacklist_file = askopenfilename(filetypes=[("Txt Files", "*.txt")], initialdir=os.path.dirname(__file__), title="Pick blacklist file")

        if self.blacklist_file:
            self.blacklist_btn.config(text=os.path.basename(self.blacklist_file))
        else:
            self.blacklist_btn.config(text="Pick File")

    def reset_file_pickers(self):
        self.tracker_file = None
        self.email_file = None
        self.email_mapping_file = None
        self.blacklist_file = None

        self.tracker_btn.config(text="Pick File")
        self.email_btn.config(text="Pick File")
        self.email_mapping_btn.config(text="Pick File")
        self.blacklist_btn.config(text="Pick File")

    def load_files(self):
        if self.blacklist_file:
            self.manager.load_email_blacklist(self.blacklist_file)

        if self.email_mapping_file:
            self.manager.load_email_mapping(self.email_mapping_file)

        if self.tracker_file:
            self.manager.load_tracker_csv(self.tracker_file)

        self.reset_file_pickers()
        # self.info_label.config(
        #     text=f"Loaded Tracker Entries: {len(self.manager.people) - 1} | Loaded Mapping Entries: {len(self.manager.email_mappings)} | Loaded Blacklist Entries: {len(self.manager.blacklist)}")

        self.tracker_count.config(text=str(len(self.manager.people) - 1))
        self.email_mapping_count.config(text=str(len(self.manager.email_mappings)))
        self.blacklist_count.config(text=str(len(self.manager.blacklist)))

    def reset_data_cb(self):
        self.manager.reset_manager()
        self.reset_file_pickers()

        self.unlock_input_buttons()
        self.lock_extract_buttons()

        self.unknown_emails = {}

    def lock_input_buttons(self):
        self.lock_input = False

        self.tracker_btn.config(text="Disabled", state="disabled")
        self.tracker_file = None

        self.email_mapping_btn.config(text="Disabled", state="disabled")
        self.email_mapping_file = None

        self.blacklist_btn.config(text="Disabled", state="disabled")
        self.blacklist_file = None

        self.load_files_btn.config(state="disabled")

    def unlock_input_buttons(self):
        self.lock_input = False

        self.tracker_btn.config(text="Pick File", state="normal")
        self.tracker_file = None

        self.email_mapping_btn.config(text="Pick File", state="normal")
        self.email_mapping_file = None

        self.blacklist_btn.config(text="Pick File", state="normal")
        self.blacklist_file = None

        self.load_files_btn.config(state="normal")

    def lock_extract_buttons(self):
        self.lock_extract = True

        self.ext_map_btn.config(state="disabled")
        self.ext_tot_btn.config(state="disabled")
        self.ext_week_btn.config(state="disabled")

    def unlock_extract_buttons(self):
        self.lock_extract = False

        self.ext_map_btn.config(state="normal")
        self.ext_tot_btn.config(state="normal")
        self.ext_week_btn.config(state="normal")

    def run_cb(self):
        if not self.email_file:
            return

        self.lock_input_buttons()

        data = data_parser.extract_emails(self.email_file)
        self.unknown_emails.update(self.manager.compile_emails(data))

        self.email_file = None
        self.email_btn.config(text="Pick File")

        self.unlock_extract_buttons()

    def ext_map_cb(self):
        if self.lock_extract:
            return

        file = asksaveasfilename(filetypes=[("JSON Files", "*.json")], initialdir=os.path.dirname(__file__), title="Pick Email Map File")
        if not file:
            print("Save file not found")
            return

        if not file.endswith(".json"):
            file += ".json"

        with open(file, "w") as f:
            json.dump(self.unknown_emails, f, indent=2)

    def ext_tot_cb(self):
        if self.lock_extract:
            return

        file = asksaveasfilename(filetypes=[("CSV Files", "*.csv")], initialdir=os.path.dirname(__file__), title="Pick Total Extract File")
        if not file:
            print("Save file not found")
            return

        if not file.endswith(".csv"):
            file += ".csv"

        total_emails = self.manager.extract_total_emails()
        with open(file, "w", newline='') as f:
            writer = csv.writer(f, delimiter=",", )
            writer.writerow(("First Name", "Preferred Name", "Last Name", "Email Count"))
            writer.writerows(total_emails)

    def ext_week_cb(self):
        if self.lock_extract:
            return

        file = asksaveasfilename(filetypes=[("Txt Files", "*.txt")], initialdir=os.path.dirname(__file__), title="Pick Weekly Extract File")
        if not file:
            print("Save file not found")
            return

        if not file.endswith(".txt"):
            file += ".txt"

        # TODO - Make this configurable
        weekly_emails = self.manager.extract_weekly_emails(date(year=2024, month=5, day=22), 12)
        with open(file, "w") as f:
            f.write("\n".join(map(str, weekly_emails)))

    def stdout(self, msg):
        self.orig_stdout.write(msg)


def main():
    root = tk.Tk()
    root.title("AR Email Tracker")
    root.geometry("500x300")
    root.resizable(False, False)

    MainFrame(root).tkraise()

    root.mainloop()


if __name__ == "__main__":
    main()
