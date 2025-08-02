""" SEO Description Optimizer GUI Application

This script provides a Tkinter-based graphical user interface for optimizing
Excel or CSV descriptions for SEO using Open AI's API. Users can upload a file
with 'Name', 'Id', and 'Description' columns, specify an SEO prompt in the first
row, and download an optimized Excel file. The interface features a modern dark
theme with vibrant accents.

Usage:
    Run the script to launch the GUI. Upload a .xlsx or .csv file, process it,
    and download the optimized output.

Dependencies:
    pandas, openpyxl, openai, tkinter
"""

import pandas as pd
from openai import OpenAI
import os
import re
from zipfile import BadZipFile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tempfile
import shutil
import threading

# Initialize OpenAI client with API key
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY", "sk-proj-..."))

# Generate an SEO-friendly version of a given description using OpenAI
def make_seo_friendly(description, seo_prompt):
    if not isinstance(description, str) or description.strip() == '':
        return "No description provided."

    prompt = f"{seo_prompt}: '{description}'"

    try:
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are an SEO expert specializing in effective and readable descriptions."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=500,
            temperature=0.7
        )
        return response.choices[0].message.content.strip().capitalize()
    except Exception as e:
        print(f"Error with Open AI API: {e}")
        return description.capitalize()

# Render a DataFrame in a Treeview widget
def display_dataframe(tree, df):
    tree.delete(*tree.get_children())
    if df is None:
        return

    columns = list(df.columns)
    tree["columns"] = columns
    tree["show"] = "headings"

    # Auto-adjust column widths
    for col in columns:
        max_width = max([len(str(df[col].iloc[i])) * 8 for i in range(min(len(df), 100))] + [len(col) * 8])
        tree.column(col, width=max(max_width, 100), anchor="w", stretch=False)
        tree.heading(col, text=col)

    for _, row in df.iterrows():
        tree.insert("", "end", values=[str(row[col]) for col in columns])

# Main application class
class SEOOptimizerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SEO Description Optimizer")
        self.root.geometry("1200x600")
        self.root.configure(bg="#1E1E2F")

        # UI setup
        self.welcome_label = tk.Label(
            root,
            text="Description Optimizer for SEO",
            font=("Helvetica", 18, "bold"),
            fg="#FFD700",
            bg="#1E1E2F",
            wraplength=1100
        )
        self.welcome_label.pack(pady=10)

        self.button_frame = tk.Frame(root, bg="#1E1E2F", bd=2, relief="sunken")
        self.button_frame.pack(fill="x", padx=10, pady=5)

        self.prompt_entry = tk.Entry(self.button_frame, font=("Helvetica", 12), width=60)
        self.prompt_entry.insert(0, "Enter custom SEO prompt here...")
        self.prompt_entry.pack(side="left", padx=5, pady=5)

        self.upload_button = tk.Button(
            self.button_frame,
            text="Upload File",
            command=self.upload_file,
            bg="white",
            fg="black",
            font=("Helvetica", 12, "bold"),
            width=15,
            activebackground="grey",
            activeforeground="black",
            relief="raised",
            cursor="hand2",
            bd=2
        )
        self.upload_button.pack(side="left", padx=5, pady=5)

        self.download_button = tk.Button(
            self.button_frame,
            text="Download Optimized File",
            command=self.download_file,
            bg="white",
            fg="black",
            font=("Helvetica", 12, "bold"),
            width=25,
            activebackground="grey",
            activeforeground="black",
            relief="raised",
            cursor="hand2",
            bd=2
        )
        self.download_button.pack(side="right", padx=5, pady=5)

        self.status_label = tk.Label(
            root,
            text="Upload a .xlsx or .csv file to optimize descriptions for SEO.",
            font=("Helvetica", 11),
            fg="#E0E0E0",
            bg="#1E1E2F",
            wraplength=1100
        )
        self.status_label.pack(pady=5)

        self.paned_window = ttk.PanedWindow(root, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill="both", expand=True, padx=10, pady=10)

        # Left pane - original data
        self.left_frame = tk.Frame(self.paned_window, bg="#1E1E2F")
        self.paned_window.add(self.left_frame, weight=1)

        self.left_label = tk.Label(
            self.left_frame,
            text="Original Spreadsheet",
            font=("Helvetica", 12, "bold"),
            fg="#FFD700",
            bg="#1E1E2F"
        )
        self.left_label.pack(pady=5)

        self.original_tree = ttk.Treeview(self.left_frame)
        self.original_tree.pack(fill="both", expand=True, padx=5, pady=5)
        self.original_scroll_y = ttk.Scrollbar(self.left_frame, orient="vertical", command=self.original_tree.yview)
        self.original_scroll_y.pack(side="right", fill="y")
        self.original_scroll_x = ttk.Scrollbar(self.left_frame, orient="horizontal", command=self.original_tree.xview)
        self.original_scroll_x.pack(side="bottom", fill="x")
        self.original_tree.configure(yscrollcommand=self.original_scroll_y.set, xscrollcommand=self.original_scroll_x.set)

        # Right pane - optimized data
        self.right_frame = tk.Frame(self.paned_window, bg="#1E1E2F")
        self.paned_window.add(self.right_frame, weight=1)

        self.right_label = tk.Label(
            self.right_frame,
            text="Optimized Spreadsheet",
            font=("Helvetica", 12, "bold"),
            fg="#FFD700",
            bg="#1E1E2F"
        )
        self.right_label.pack(pady=5)

        self.optimized_tree = ttk.Treeview(self.right_frame)
        self.optimized_tree.pack(fill="both", expand=True, padx=5, pady=5)
        self.optimized_scroll_y = ttk.Scrollbar(self.right_frame, orient="vertical", command=self.optimized_tree.yview)
        self.optimized_scroll_y.pack(side="right", fill="y")
        self.optimized_scroll_x = ttk.Scrollbar(self.right_frame, orient="horizontal", command=self.optimized_tree.xview)
        self.optimized_scroll_x.pack(side="bottom", fill="x")
        self.optimized_tree.configure(yscrollcommand=self.optimized_scroll_y.set, xscrollcommand=self.optimized_scroll_x.set)

        # Internal state variables
        self.input_path = None
        self.output_path = None
        self.optimized_df = None
        self.original_df = None

    # Load and display file
    def upload_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel/CSV files", "*.xlsx *.csv")]
        )
        if not file_path:
            return

        self.input_path = file_path
        self.status_label.config(text="Processing...", fg="#FFD700")
        self.download_button.config(state="disabled")
        self.root.update()

        display_dataframe(self.original_tree, None)
        display_dataframe(self.optimized_tree, None)

        try:
            if file_path.endswith('.csv'):
                try:
                    original_df = pd.read_csv(file_path, encoding='utf-8')
                except UnicodeDecodeError:
                    original_df = pd.read_csv(file_path, encoding='cp1252')
            else:
                original_df = pd.read_excel(file_path, engine='openpyxl')

            display_dataframe(self.original_tree, original_df)
            self.original_df = original_df.copy()

            # Run optimization in a separate thread to keep UI responsive
            threading.Thread(target=self.optimize_seo_descriptions).start()

        except Exception as e:
            self.status_label.config(text="Failed to load original file.", fg="#FF4040")
            messagebox.showerror("Error", f"Error loading file: {e}", parent=self.root)

    # Optimize the descriptions column for SEO
    def optimize_seo_descriptions(self):
        try:
            df = self.original_df.copy()

            if 'Description' not in df.columns or 'Name' not in df.columns:
                raise ValueError("Missing required columns: 'Name' and 'Description'")

            seo_prompt = self.prompt_entry.get().strip() or df.loc[0, 'Name'].strip()
            df = df.drop(index=0).reset_index(drop=True)

            for idx in df.index:
                desc = df.at[idx, 'Description']
                optimized = make_seo_friendly(desc, seo_prompt)
                df.at[idx, 'Description'] = optimized
                display_dataframe(self.optimized_tree, df.iloc[:idx+1])

            self.optimized_df = df
            self.status_label.config(text="Optimization complete!", fg="#00FF00")
            self.download_button.config(state="normal")

        except Exception as e:
            self.status_label.config(text="Optimization failed.", fg="#FF4040")
            messagebox.showerror("Error", f"Error optimizing data: {e}", parent=self.root)

    # Save optimized data to a file
    def download_file(self):
        if not self.optimized_df is not None:
            messagebox.showerror("Error", "No optimized file available.", parent=self.root)
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="seo_optimized_data.xlsx"
        )
        if save_path:
            self.optimized_df.to_excel(save_path, index=False, engine='openpyxl')
            messagebox.showinfo("Success", f"File saved to {save_path}", parent=self.root)

# Launch the application
if __name__ == "__main__":
    root = tk.Tk()
    app = SEOOptimizerApp(root)
    root.mainloop()
