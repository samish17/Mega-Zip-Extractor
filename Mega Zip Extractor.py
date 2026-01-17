import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import zipfile
import os
from pathlib import Path
import threading


class ZipExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Multi-Zip Extractor")
        self.root.geometry("700x500")
        self.root.resizable(True, True)
        
        # Variables
        self.zip_files = []
        self.output_dir = tk.StringVar()
        self.extract_to_subfolders = tk.BooleanVar(value=True)
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Multi-Zip Extractor", 
                                font=("Arial", 16, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=1, column=0, columnspan=3, pady=(0, 10), sticky=(tk.W, tk.E))
        
        ttk.Button(buttons_frame, text="Select Zip Files", 
                   command=self.select_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(buttons_frame, text="Select Folder with Zips", 
                   command=self.select_folder).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Clear List", 
                   command=self.clear_list).pack(side=tk.LEFT, padx=5)
        
        # Files listbox with scrollbar
        list_frame = ttk.LabelFrame(main_frame, text="Selected Zip Files", padding="5")
        list_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        self.files_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, height=10)
        self.files_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.config(command=self.files_listbox.yview)
        
        # Output directory
        output_frame = ttk.LabelFrame(main_frame, text="Output Directory", padding="5")
        output_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        output_frame.columnconfigure(0, weight=1)
        
        ttk.Entry(output_frame, textvariable=self.output_dir, state="readonly").grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(output_frame, text="Browse", 
                   command=self.select_output_dir).grid(row=0, column=1)
        
        # Options
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="5")
        options_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Checkbutton(options_frame, 
                        text="Extract each zip to a separate subfolder (recommended)",
                        variable=self.extract_to_subfolders).pack(anchor=tk.W)
        
        # Progress
        progress_frame = ttk.Frame(main_frame)
        progress_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress_label = ttk.Label(progress_frame, text="Ready")
        self.progress_label.grid(row=0, column=0, sticky=tk.W)
        
        self.progress_bar = ttk.Progressbar(progress_frame, mode='determinate')
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
        # Extract button
        self.extract_button = ttk.Button(main_frame, text="Extract All", 
                                         command=self.extract_all, style="Accent.TButton")
        self.extract_button.grid(row=6, column=0, columnspan=3, pady=(10, 0))
        
    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Select Zip Files",
            filetypes=[("Zip files", "*.zip"), ("All files", "*.*")]
        )
        if files:
            for file in files:
                if file not in self.zip_files:
                    self.zip_files.append(file)
                    self.files_listbox.insert(tk.END, os.path.basename(file))
            self.update_status()
    
    def select_folder(self):
        folder = filedialog.askdirectory(title="Select Folder Containing Zip Files")
        if folder:
            zip_files = list(Path(folder).glob("*.zip"))
            if zip_files:
                for zip_file in zip_files:
                    zip_path = str(zip_file)
                    if zip_path not in self.zip_files:
                        self.zip_files.append(zip_path)
                        self.files_listbox.insert(tk.END, zip_file.name)
                self.update_status()
            else:
                messagebox.showinfo("No Zip Files", "No zip files found in the selected folder.")
    
    def select_output_dir(self):
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_dir.set(directory)
    
    def clear_list(self):
        self.zip_files.clear()
        self.files_listbox.delete(0, tk.END)
        self.update_status()
    
    def update_status(self):
        count = len(self.zip_files)
        self.progress_label.config(text=f"Ready - {count} file(s) selected")
    
    def extract_all(self):
        if not self.zip_files:
            messagebox.showwarning("No Files", "Please select zip files to extract.")
            return
        
        output_dir = self.output_dir.get()
        if not output_dir:
            messagebox.showwarning("No Output Directory", 
                                   "Please select an output directory.")
            return
        
        # Run extraction in a separate thread to keep UI responsive
        thread = threading.Thread(target=self.perform_extraction, daemon=True)
        thread.start()
    
    def perform_extraction(self):
        output_dir = self.output_dir.get()
        total_files = len(self.zip_files)
        
        self.extract_button.config(state="disabled")
        self.progress_bar["maximum"] = total_files
        self.progress_bar["value"] = 0
        
        extracted = 0
        failed = []
        
        for i, zip_path in enumerate(self.zip_files, 1):
            try:
                zip_name = Path(zip_path).stem
                self.progress_label.config(text=f"Extracting {i}/{total_files}: {os.path.basename(zip_path)}")
                
                if self.extract_to_subfolders.get():
                    extract_path = os.path.join(output_dir, zip_name)
                else:
                    extract_path = output_dir
                
                os.makedirs(extract_path, exist_ok=True)
                
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    zip_ref.extractall(extract_path)
                
                extracted += 1
                self.progress_bar["value"] = i
                
            except Exception as e:
                failed.append(f"{os.path.basename(zip_path)}: {str(e)}")
        
        # Show completion message
        self.root.after(0, lambda: self.show_completion(extracted, failed))
        self.root.after(0, lambda: self.extract_button.config(state="normal"))
    
    def show_completion(self, extracted, failed):
        self.progress_label.config(text=f"Completed - {extracted} file(s) extracted")
        
        if failed:
            msg = f"Extraction completed with errors:\n\n"
            msg += f"Successfully extracted: {extracted}\n"
            msg += f"Failed: {len(failed)}\n\n"
            msg += "Failed files:\n" + "\n".join(failed[:5])
            if len(failed) > 5:
                msg += f"\n... and {len(failed) - 5} more"
            messagebox.showwarning("Extraction Complete with Errors", msg)
        else:
            messagebox.showinfo("Success", 
                                f"All {extracted} zip files extracted successfully!")


def main():
    root = tk.Tk()
    app = ZipExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()