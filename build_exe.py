import PyInstaller.__main__
import os
import customtkinter

# Define the script to build
script_name = "gui_app.py"
app_name = "ExcelSplitter"

# PyInstaller arguments
args = [
    script_name,
    f"--name={app_name}",
    "--onefile",           # Bundle into a single executable
    "--noconsole",         # Hide the console window
    "--clean",             # Clean cache
    # Add hidden imports if necessary (pandas often needs these)
    "--hidden-import=pandas",
    "--hidden-import=openpyxl",
    "--hidden-import=customtkinter",
    "--hidden-import=winsound",
    # Add customtkinter data files (themes, fonts, etc.)
    f"--add-data={os.path.dirname(customtkinter.__file__)};customtkinter",
]

print("Building executable...")
PyInstaller.__main__.run(args)
print(f"Build complete. Check the 'dist' folder for {app_name}.exe")
