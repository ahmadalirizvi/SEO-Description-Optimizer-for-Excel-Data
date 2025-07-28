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

# Initialize Open AI client
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY", "key!"))  # Use environment variable in production


def make_seo_friendly(description: str, seo_prompt: str) -> str:
    """
    Optimize a description for SEO using Open AI's API.

    Args:
        description (str): The original description to optimize.
        seo_prompt (str): The SEO prompt to guide optimization.

    Returns:
        str: The optimized description, capitalized. Returns original description
             if API call fails or input is invalid.
    """
    if not isinstance(description, str) or description.strip() == '':
        return "No description provided."

    prompt = f"{seo_prompt}: '{description}'"

    try:
        # Call Open AI API to generate optimized description
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are an SEO expert specializing in effective and readable descriptions."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=500,
            temperature=0.7
        )
        optimized_description = response.choices[0].message.content.strip()
        return optimized_description.capitalize()

    except Exception as e:
        print(f"Error with Open AI API: {e}")
        return description.capitalize()


def process_excel_file(input_path: str, output_path: str) -> tuple[bool, str]:
    """
    Process an Excel or CSV file to optimize descriptions for SEO.

    Args:
        input_path (str): Path to the input .xlsx or .csv file.
        output_path (str): Path to save the optimized .xlsx file.

    Returns:
        tuple[bool, str]: A tuple containing a success flag and a status message.
    """
    try:
        # Read input file based on extension
        if input_path.endswith('.csv'):
            df = pd.read_csv(input_path)
        else:
            try:
                df = pd.read_excel(input_path, engine='openpyxl')
            except BadZipFile:
                raise ValueError("The file is not a valid .xlsx file.")

        # Validate required columns
        required_columns = ['Name', 'Id', 'Description']
        if not all(col in df.columns for col in required_columns):
            raise ValueError("Excel file must contain 'Name', 'Id', and 'Description' columns.")

        # Validate SEO prompt in first row
        if len(df) < 1 or not isinstance(df.loc[0, 'Name'], str) or df.loc[0, 'Name'].strip() == '':
            raise ValueError("The first cell in the 'Name' column must contain a valid SEO prompt.")
        seo_prompt = df.loc[0, 'Name'].strip()

        # Optimize descriptions for rows after the first
        df.loc[1:, 'Description'] = df.loc[1:, 'Description'].apply(lambda desc: make_seo_friendly(desc, seo_prompt))

        # Save optimized data to Excel
        df.to_excel(output_path, index=False, engine='openpyxl')
        return True, "SEO-friendly Excel file saved."

    except ValueError as e:
        return False, f"Error: {e}"
    except Exception as e:
        return False, f"An unexpected error occurred: {e}"


class SEOOptimizerApp:
    """A Tkinter-based GUI application for SEO description optimization."""

    def __init__(self, root: tk.Tk):
        """
        Initialize the GUI application.

        Args:
            root (tk.Tk): The root Tkinter window.
        """
        self.root = root
        self.root.title("SEO Description Optimizer")
        self.root.geometry("600x400")
        self.root.configure(bg="#1E1E2F")  # Dark blue-gray background

        # Initialize GUI components
        self._setup_gui()

        self.input_path = None
        self.output_path = None

    def _setup_gui(self):
        """Configure GUI elements for the application."""
        # Welcome message
        self.welcome_label = tk.Label(
            self.root,
            text="Welcome to SEO Description Optimizer!",
            font=("Helvetica", 18, "bold"),
            fg="#FFD700",  # Gold text
            bg="#1E1E2F",
            wraplength=500
        )
        self.welcome_label.pack(pady=20)

        # Instructions
        self.instructions = tk.Label(
            self.root,
            text="Upload a .xlsx or .csv file to optimize descriptions for SEO.",
            font=("Helvetica", 11),
            fg="#E0E0E0",  # Light gray
            bg="#1E1E2F"
        )
        self.instructions.pack(pady=10)

        # File upload button
        self.upload_button = tk.Button(
            self.root,
            text="Upload File",
            command=self.upload_file,
            bg="#00A896",  # Teal
            fg="white",
            font=("Helvetica", 12, "bold"),
            width=20,
            activebackground="#02C39A",  # Lighter teal on hover
            activeforeground="white",
            relief="flat",
            cursor="hand2"
        )
        self.upload_button.pack(pady=10)

        # Status label
        self.status_label = tk.Label(
            self.root,
            text="",
            font=("Helvetica", 11),
            fg="#E0E0E0",
            bg="#1E1E2F",
            wraplength=500
        )
        self.status_label.pack(pady=10)

        # Download button (initially disabled)
        self.download_button = tk.Button(
            self.root,
            text="Download Optimized File",
            command=self.download_file,
            bg="#FF6B6B",  # Coral
            fg="white",
            font=("Helvetica", 12, "bold"),
            width=20,
            state="disabled",
            activebackground="#FF8787",  # Lighter coral on hover
            activeforeground="white",
            relief="flat",
            cursor="hand2"
        )
        self.download_button.pack(pady=10)

    def upload_file(self):
        """Handle file upload and processing."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel/CSV files", "*.xlsx *.csv")]
        )
        if not file_path:
            return

        self.input_path = file_path
        self.status_label.config(text="Processing...", fg="#FFD700")
        self.download_button.config(state="disabled")
        self.root.update()

        # Create temporary output file
        self.output_path = os.path.join(tempfile.gettempdir(), "seo_optimized_data.xlsx")

        # Process the file and update UI
        success, message = process_excel_file(self.input_path, self.output_path)
        if success:
            self.status_label.config(text="Processing complete!", fg="#00FF00")
            self.download_button.config(state="normal")
            messagebox.showinfo("Success", message, parent=self.root)
        else:
            self.status_label.config(text="Processing failed.", fg="#FF4040")
            messagebox.showerror("Error", message, parent=self.root)

    def download_file(self):
        """Handle downloading of the optimized file."""
        if not self.output_path or not os.path.exists(self.output_path):
            messagebox.showerror("Error", "No optimized file available.", parent=self.root)
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="seo_optimized_data.xlsx"
        )
        if save_path:
            shutil.copy(self.output_path, save_path)
            messagebox.showinfo("Success", f"File saved to {save_path}", parent=self.root)


if __name__ == "__main__":
    """Entry point for the application."""
    root = tk.Tk()
    app = SEOOptimizerApp(root)
    root.mainloop()
