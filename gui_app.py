import customtkinter as ctk
import sys
import os
import threading
import io
import contextlib
from tkinter import filedialog, messagebox
import winsound
import account  # Importing the original script

# Set theme
ctk.set_appearance_mode("Light")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class ConsoleRedirector:
    """Redirects stdout to a text widget."""
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, str):
        self.text_widget.configure(state="normal")
        self.text_widget.insert("end", str)
        self.text_widget.see("end")
        self.text_widget.configure(state="disabled")

    def flush(self):
        pass

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window configuration
        self.title("Excel Account Splitter")
        self.geometry("700x500")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        # Header
        self.header_frame = ctk.CTkFrame(self, corner_radius=0)
        self.header_frame.grid(row=0, column=0, sticky="ew")
        self.header_label = ctk.CTkLabel(self.header_frame, text="Excel Account Splitter", font=ctk.CTkFont(size=20, weight="bold"))
        self.header_label.pack(padx=20, pady=15)

        # File Selection
        self.file_frame = ctk.CTkFrame(self)
        self.file_frame.grid(row=1, column=0, padx=20, pady=20, sticky="ew")
        self.file_frame.grid_columnconfigure(0, weight=1)

        self.file_entry = ctk.CTkEntry(self.file_frame, placeholder_text="Select Excel File...")
        self.file_entry.grid(row=0, column=0, padx=(20, 10), pady=20, sticky="ew")

        self.browse_button = ctk.CTkButton(self.file_frame, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=0, column=1, padx=(0, 20), pady=20)

        # Action Buttons
        self.action_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.action_frame.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="ew")
        
        self.process_button = ctk.CTkButton(self.action_frame, text="Process File", command=self.start_processing, height=40, font=ctk.CTkFont(size=15, weight="bold"))
        self.process_button.pack(fill="x")

        # Log Area
        self.log_textbox = ctk.CTkTextbox(self, width=660, height=200)
        self.log_textbox.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="nsew")
        self.log_textbox.configure(state="disabled")

        # Output Button (Initially Hidden)
        self.open_output_button = ctk.CTkButton(self.action_frame, text="Open Output Folder", command=self.open_output_folder, fg_color="green", height=0)
        # It will be packed when needed

        self.selected_file = None

    def browse_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if filename:
            self.selected_file = filename
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, filename)

    def start_processing(self):
        if not self.selected_file:
            messagebox.showerror("Error", "Please select a file first.")
            return

        self.process_button.configure(state="disabled", text="Processing...")
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")

        # Run in a separate thread to keep GUI responsive
        threading.Thread(target=self.run_script, daemon=True).start()

    def run_script(self):
        # Redirect stdout/stderr
        original_stdout = sys.stdout
        original_stderr = sys.stderr
        sys.stdout = ConsoleRedirector(self.log_textbox)
        sys.stderr = ConsoleRedirector(self.log_textbox)

        try:
            # Mock sys.argv
            sys.argv = ["account.py", self.selected_file]
            
            print(f"Starting process for: {self.selected_file}...\n")
            
            # Call the main function of the imported script
            account.main()
            
            print("\nDone.")
            # Play success sound (System Default)
            winsound.MessageBeep(winsound.MB_OK)

        except Exception as e:
            print(f"\n❌ Error: {str(e)}")
            self.after(0, lambda: messagebox.showerror("Error", f"An error occurred:\n{str(e)}"))
        
        finally:
            # Restore stdout/stderr
            sys.stdout = original_stdout
            sys.stderr = original_stderr
            self.after(0, self.reset_ui)

    def reset_ui(self):
        self.process_button.configure(state="normal", text="Process File")
        self.show_output_button()

    def show_output_button(self):
        # Check if already packed to avoid double packing
        if not self.open_output_button.winfo_ismapped():
            self.open_output_button.pack(fill="x", pady=(10, 0))
            self.animate_button_height(0)

    def animate_button_height(self, current_height):
        target_height = 40
        if current_height < target_height:
            current_height += 2
            self.open_output_button.configure(height=current_height)
            self.after(10, lambda: self.animate_button_height(current_height))
        else:
            self.open_output_button.configure(height=target_height)

    def open_output_folder(self):
        if self.selected_file:
            folder = os.path.dirname(self.selected_file)
            os.startfile(folder)
        else:
            # Default to current directory if no file selected
            os.startfile(os.getcwd())

if __name__ == "__main__":
    app = App()
    app.mainloop()
