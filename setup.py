import sys
from cx_Freeze import setup, Executable

# Dependencies are automatically detected, but may need fine-tuning.
build_exe_options = {
    "packages": ["os", "tkinter", "PIL", "docx", "lxml", "babel", "pytz", "pkg_resources", "platform", "xml",
                 "sysconfig", "darkdetect",
                 "difflib", "multiprocessing", "zoneinfo", "jinja2", "sqlite3", "customtkinter"],
    "excludes": [],
    "include_files": [
        ("images", "images"),  # Include the entire 'images' directory
        ("logo/logo_icon.ico", "logo/logo_icon.ico"),  # Include the icon
        ("logo", "logo"),
        ("template", "template"),
        # ("path/to/other_file.txt", "other_file.txt"), # Include individual files
    ],
}

# base="Win32GUI" should be used only for Windows GUI app
base = "Win32GUI" if sys.platform == "win32" else None

setup(
    name="SmartDoc",
    version="1.0",
    description="Make Docx easier for doctors ",
    options={"build_exe": build_exe_options},
    executables=[Executable("main.py", base=base, target_name="SmartDoc.exe", icon="logo/logo_icon.ico")]
)
