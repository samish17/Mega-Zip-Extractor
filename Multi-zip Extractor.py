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
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Set color scheme
        self.bg_color = "#1e1e2e"
        self.secondary_bg = "#2a2a3e"
        self.accent_color = "#6366f1"
        self.accent_hover = "#4f46e5"
        self.text_color = "#e0e0e0"
        self.success_color = "#10b981"
        self.warning_color = "#f59e0b"
        
        self.root.configure(bg=self.bg_color)
        
        # Variables
        self.zip_files = []
        self.output_dir = tk.StringVar()
        self.extract_to_subfolders = tk.BooleanVar(value=True)
        
        self.setup_styles()
        self.setup_ui()
        
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure colors
        style.configure(".", background=self.bg_color, foreground=self.text_color,
                       fieldbackground=self.secondary_bg, borderwidth=0)
        
        # Frame styles
        style.configure("Card.TFrame", background=self.secondary_bg, relief="flat")
        style.configure("Main.TFrame", background=self.bg_color)
        
        # Label styles
        style.configure("Title.TLabel", background=self.bg_color, foreground=self.text_color,
                       font=("Segoe UI", 24, "bold"))
        style.configure("Subtitle.TLabel", background=self.bg_color, foreground="#a0a0b0",
                       font=("Segoe UI", 10))
        style.configure("Card.TLabel", background=self.secondary_bg, foreground=self.text_color,
                       font=("Segoe UI", 10))
        style.configure("Status.TLabel", background=self.bg_color, foreground="#a0a0b0",
                       font=("Segoe UI", 9))
        
        # Button styles
        style.configure("Accent.TButton", background=self.accent_color, foreground="white",
                       borderwidth=0, focuscolor="none", font=("Segoe UI", 10, "bold"),
                       padding=(20, 10))
        style.map("Accent.TButton",
                 background=[("active", self.accent_hover), ("pressed", "#3730a3")])
        
        style.configure("Secondary.TButton", background=self.secondary_bg, 
                       foreground=self.text_color, borderwidth=0, focuscolor="none",
                       font=("Segoe UI", 9), padding=(15, 8))
        style.map("Secondary.TButton",
                 background=[("active", "#35354f"), ("pressed", "#2a2a3e")])
        
        # Checkbutton style
        style.configure("Card.TCheckbutton", background=self.secondary_bg,
                       foreground=self.text_color, font=("Segoe UI", 9))
        
        # Progress bar style
        style.configure("Horizontal.TProgressbar", background=self.accent_color,
                       troughcolor=self.secondary_bg, borderwidth=0, thickness=8)
        
        # LabelFrame style
        style.configure("Card.TLabelframe", background=self.secondary_bg,
                       foreground=self.text_color, borderwidth=0, relief="flat")
        style.configure("Card.TLabelframe.Label", background=self.secondary_bg,
                       foreground=self.text_color, font=("Segoe UI", 10, "bold"))
        
    def setup_ui(self):
        # Main container with padding
        container = ttk.Frame(self.root, style="Main.TFrame", padding="30")
        container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(3, weight=1)
        
        # Header
        header_frame = ttk.Frame(container, style="Main.TFrame")
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 30))
        
        title_label = ttk.Label(header_frame, text="📦 Zip Extractor", 
                                style="Title.TLabel")
        title_label.pack(anchor=tk.W)
        
        subtitle_label = ttk.Label(header_frame, 
                                   text="Extract multiple zip files with ease",
                                   style="Subtitle.TLabel")
        subtitle_label.pack(anchor=tk.W, pady=(5, 0))
        
        # Action buttons card
        buttons_card = ttk.Frame(container, style="Card.TFrame", padding="20")
        buttons_card.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        
        buttons_inner = ttk.Frame(buttons_card, style="Card.TFrame")
        buttons_inner.pack(fill=tk.X)
        
        ttk.Button(buttons_inner, text="📁 Select Files", 
                   command=self.select_files, style="Secondary.TButton").pack(
                       side=tk.LEFT, padx=(0, 10))
        ttk.Button(buttons_inner, text="📂 Select Folder", 
                   command=self.select_folder, style="Secondary.TButton").pack(
                       side=tk.LEFT, padx=(0, 10))
        ttk.Button(buttons_inner, text="🗑️ Clear", 
                   command=self.clear_list, style="Secondary.TButton").pack(
                       side=tk.LEFT)
        
        # Files list card
        files_card = ttk.Frame(container, style="Card.TFrame", padding="20")
        files_card.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        
        files_label = ttk.Label(files_card, text="Selected Files", style="Card.TLabel")
        files_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Custom listbox
        list_frame = ttk.Frame(files_card, style="Card.TFrame")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(list_frame, bg=self.secondary_bg, 
                                troughcolor=self.secondary_bg, bd=0, 
                                highlightthickness=0, activebackground=self.accent_color)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.files_listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set, 
                                        height=8, bg=self.bg_color, fg=self.text_color,
                                        selectbackground=self.accent_color,
                                        selectforeground="white", bd=0,
                                        highlightthickness=1, highlightcolor=self.accent_color,
                                        highlightbackground="#35354f",
                                        font=("Segoe UI", 9))
        self.files_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.files_listbox.yview)
        
        # Output directory card
        output_card = ttk.Frame(container, style="Card.TFrame", padding="20")
        output_card.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N), pady=(0, 20))
        
        output_label = ttk.Label(output_card, text="Output Directory", style="Card.TLabel")
        output_label.pack(anchor=tk.W, pady=(0, 10))
        
        output_inner = ttk.Frame(output_card, style="Card.TFrame")
        output_inner.pack(fill=tk.X)
        
        output_entry = tk.Entry(output_inner, textvariable=self.output_dir, 
                               state="readonly", bg=self.bg_color, fg=self.text_color,
                               insertbackground=self.text_color, bd=0,
                               highlightthickness=1, highlightcolor=self.accent_color,
                               highlightbackground="#35354f", font=("Segoe UI", 9))
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10), ipady=8)
        
        ttk.Button(output_inner, text="Browse", command=self.select_output_dir,
                   style="Secondary.TButton").pack(side=tk.LEFT)
        
        # Options
        options_check = ttk.Checkbutton(output_card, 
                        text="Extract each zip to a separate subfolder",
                        variable=self.extract_to_subfolders,
                        style="Card.TCheckbutton")
        options_check.pack(anchor=tk.W, pady=(15, 0))
        
        # Progress section
        progress_card = ttk.Frame(container, style="Card.TFrame", padding="20")
        progress_card.grid(row=4, column=0, sticky=(tk.W, tk.E, tk.N), pady=(0, 20))
        
        self.progress_label = ttk.Label(progress_card, text="Ready to extract", 
                                       style="Status.TLabel")
        self.progress_label.pack(anchor=tk.W, pady=(0, 10))
        
        self.progress_bar = ttk.Progressbar(progress_card, mode='determinate',
                                           style="Horizontal.TProgressbar")
        self.progress_bar.pack(fill=tk.X)
        
        self.file_count_label = ttk.Label(progress_card, text="0 files selected",
                                         style="Status.TLabel")
        self.file_count_label.pack(anchor=tk.W, pady=(8, 0))
        
        # Extract button
        button_frame = ttk.Frame(container, style="Main.TFrame")
        button_frame.grid(row=5, column=0, sticky=(tk.W, tk.E))
        
        self.extract_button = ttk.Button(button_frame, text="✨ Extract All Files", 
                                         command=self.extract_all, 
                                         style="Accent.TButton")
        self.extract_button.pack(fill=tk.X)
        
    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Select Zip Files",
            filetypes=[("Zip files", "*.zip"), ("All files", "*.*")]
        )
        if files:
            for file in files:
                if file not in self.zip_files:
                    self.zip_files.append(file)
                    self.files_listbox.insert(tk.END, f"  📄 {os.path.basename(file)}")
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
                        self.files_listbox.insert(tk.END, f"  📄 {zip_file.name}")
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
        plural = "file" if count == 1 else "files"
        self.file_count_label.config(text=f"{count} {plural} selected")
        self.progress_label.config(text="Ready to extract")
    
    def extract_all(self):
        if not self.zip_files:
            messagebox.showwarning("No Files", "Please select zip files to extract.")
            return
        
        output_dir = self.output_dir.get()
        if not output_dir:
            messagebox.showwarning("No Output Directory", 
                                   "Please select an output directory.")
            return
        
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
                self.progress_label.config(
                    text=f"⚡ Extracting {i} of {total_files}: {os.path.basename(zip_path)}")
                
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
        
        self.root.after(0, lambda: self.show_completion(extracted, failed))
        self.root.after(0, lambda: self.extract_button.config(state="normal"))
    
    def show_completion(self, extracted, failed):
        self.progress_label.config(text=f"✅ Completed - {extracted} file(s) extracted successfully")
        
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
                                f"🎉 All {extracted} zip files extracted successfully!")


def main():
    root = tk.Tk()
    app = ZipExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()