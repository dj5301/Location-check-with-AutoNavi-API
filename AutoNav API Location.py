import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import requests
import time
import threading

# GUITools: Gaode Fuzzy Query Desktop Version

class GaodeApp:
    def __init__(self, master):
        self.master = master
        master.title("📍 Location Check")
        master.geometry("550x280")

        # Load Excel file button
        tk.Button(master, text="Choose the excel that contain the data", command=self.load_excel).pack(pady=10)
        self.status_label = tk.Label(master, text="📂 No file selected")
        self.status_label.pack()

        # Begin searching button
        self.query_button = tk.Button(master, text="Begin searching", command=self.run_query_thread, bg="lightgreen")
        self.query_button.pack(pady=15)

        # Keyword list
        self.keywords = []

    def load_excel(self):
        path = filedialog.askopenfilename(title="Choose your excel", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if path:
            try:
                df = pd.read_excel(path)
                self.keywords = df.iloc[:, 0].dropna().astype(str).tolist()
                self.status_label.config(text=f"✅ Processing {len(self.keywords)} keywords")
            except Exception as e:
                messagebox.showerror("Failed", f"Failed to read the document：{e}")
        else:
            self.status_label.config(text="❌ Cancelled")

    def query_address(self, keyword, key):
        url = "https://restapi.amap.com/v3/geocode/geo"
        params = {"address": keyword, "key": key, "output": "JSON"}
        try:
            r = requests.get(url, params=params, timeout=10)
            data = r.json()
            if data["status"] == "1" and int(data["count"]) > 0:
                geo = data["geocodes"][0]
                return {
                    "Keyword": keyword,
                    "Location": geo.get("formatted_address", ""),
                    "Province/State": geo.get("province", ""),
                    "City": geo.get("city", ""),
                    "County": geo.get("district", ""),
                    "Longitute & Latitude": geo.get("location", "")
                }
        except Exception as e:
            return {"Keywords": keyword, "Location": f"Failed: {e}", "Province/State": "", "City": "", "County": "", "Longitute & Latitude": ""}
        return {"Keywords": keyword, "Location": "No result", "Province/State": "", "City": "", "County": "", "Longitute & Latitude": ""}

    def run_query_thread(self):
        thread = threading.Thread(target=self.run_query)
        thread.start()

    def run_query(self):
        key = "Your key" # (Substitute with actual key)
        if not key:
            messagebox.showwarning("Attention", "Plse fill in the API Key first")
            return
        if not self.keywords:
            messagebox.showwarning("Attention", "Plse load the excel file first")
            return

        self.query_button.config(state="disabled")
        results = []
        for i, kw in enumerate(self.keywords):
            self.status_label.config(text=f"🔍 Searching ({i+1}/{len(self.keywords)})：{kw}")
            results.append(self.query_address(kw, key))
            time.sleep(0.3)
        df = pd.DataFrame(results)
        

# Pop up save path dialog
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel file", "*.xlsx")],
            title="Save the result as..."
        )

        if save_path:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Finished", f"✅ Finished and save your result at：\n{save_path}")
        else:
            messagebox.showinfo("Cancel", "❌ User cancelled")

        self.status_label.config(text="✅ Finished")
        self.query_button.config(state="normal")


# Initialize and run the app
root = tk.Tk()
app = GaodeApp(root)
root.mainloop()