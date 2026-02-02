import tkinter as tk
from tkinter import messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
import os
from pathlib import Path
import threading
from excel_processor import process_excel_file


class KrecTrecApp:
    def __init__(self, root):
        self.root = root
        self.root.title("KREC/TREC Processor")
        self.root.geometry("600x400")
        self.root.configure(bg="#f0f0f0")
        
        # Center the window
        self.center_window()
        
        # Create main frame
        main_frame = tk.Frame(root, bg="#f0f0f0")
        main_frame.pack(expand=True, fill="both", padx=20, pady=20)
        
        # Title label
        title_label = tk.Label(
            main_frame,
            text="Excel File Processor",
            font=("Arial", 20, "bold"),
            bg="#f0f0f0",
            fg="#333333"
        )
        title_label.pack(pady=(0, 20))
        
        # Drop area
        self.drop_frame = tk.Frame(
            main_frame,
            bg="#ffffff",
            relief="ridge",
            borderwidth=3
        )
        self.drop_frame.pack(expand=True, fill="both", padx=20, pady=20)
        
        # Drop label
        self.drop_label = tk.Label(
            self.drop_frame,
            text="Drag and Drop Excel File Here\n\n(or click to browse)",
            font=("Arial", 16),
            bg="#ffffff",
            fg="#666666",
            justify="center"
        )
        self.drop_label.pack(expand=True)
        
        # Bind drag and drop
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
        
        # Bind click to browse
        self.drop_frame.bind("<Button-1>", self.browse_file)
        self.drop_label.bind("<Button-1>", self.browse_file)
        
        # Status label
        self.status_label = tk.Label(
            main_frame,
            text="Ready to process files",
            font=("Arial", 10),
            bg="#f0f0f0",
            fg="#666666"
        )
        self.status_label.pack(pady=(0, 10))
        
        # Processing flag
        self.is_processing = False
    
    def center_window(self):
        """Center the window on the screen"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def browse_file(self, event):
        """Open file browser"""
        from tkinter import filedialog
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.process_file(file_path)
    
    def on_drop(self, event):
        """Handle file drop event"""
        # Get the file path
        file_path = event.data
        
        # Remove curly braces if present (Windows drag-drop adds them)
        if file_path.startswith('{') and file_path.endswith('}'):
            file_path = file_path[1:-1]
        
        # Check if it's an Excel file
        if file_path.lower().endswith(('.xlsx', '.xls')):
            self.process_file(file_path)
        else:
            messagebox.showerror("Error", "Please drop an Excel file (.xlsx or .xls)")
    
    def process_file(self, file_path):
        """Process the Excel file in a separate thread"""
        if self.is_processing:
            messagebox.showwarning("Processing", "Already processing a file. Please wait.")
            return
        
        self.is_processing = True
        self.update_status("Processing file...")
        self.drop_label.config(text="Processing...\nPlease wait", fg="#ff6600")
        
        # Process in a separate thread to keep UI responsive
        thread = threading.Thread(target=self.process_file_thread, args=(file_path,))
        thread.daemon = True
        thread.start()
    
    def process_file_thread(self, file_path):
        """Process file in background thread"""
        try:
            output_path = process_excel_file(file_path)
            
            # Update UI from main thread
            self.root.after(0, self.processing_complete, output_path)
        except Exception as e:
            # Show error from main thread
            self.root.after(0, self.processing_error, str(e))
    
    def processing_complete(self, output_path):
        """Called when processing is complete"""
        self.is_processing = False
        self.drop_label.config(
            text="Drag and Drop Excel File Here\n\n(or click to browse)",
            fg="#666666"
        )
        self.update_status("Ready to process files")
        
        # Show success message with option to open folder
        result = messagebox.askyesno(
            "Success",
            f"File processed successfully!\n\nOutput saved to:\n{output_path}\n\nDo you want to open the folder?",
            icon="info"
        )
        
        if result:
            # Open folder containing the file
            folder_path = str(Path(output_path).parent)
            os.startfile(folder_path)
    
    def processing_error(self, error_message):
        """Called when processing encounters an error"""
        self.is_processing = False
        self.drop_label.config(
            text="Drag and Drop Excel File Here\n\n(or click to browse)",
            fg="#666666"
        )
        self.update_status("Ready to process files")
        
        messagebox.showerror("Error", f"Failed to process file:\n\n{error_message}")
    
    def update_status(self, message):
        """Update status label"""
        self.status_label.config(text=message)


def main():
    # Create the main window with drag-and-drop support
    root = TkinterDnD.Tk()
    app = KrecTrecApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
