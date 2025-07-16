import customtkinter as ctk
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import os
import sys
import shutil
import subprocess
import threading
from datetime import datetime
from typing import Dict, List, Any, Optional, Callable
from PIL import Image as PILImage

# Application version information
APP_VERSION = "1.0.2"
APP_BUILD_DATE = "2025-07-16"
UPDATE_SERVER_PATH = r"\\CDIMANQ30\Creoman-Active\CADCAM\Software\Primus Dental Implant Report Generator"
UPDATE_FILENAME = "Primus Dental Implant Report Generator.exe"

# Set appearance mode and color theme
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")
# Pantone Color Definitions
PANTONE_3597C = "#00B5D8"  # Light blue
# Pantone 7683C - Dark Blue
PANTONE_7683C = "#1E3A8A"  # Dark blue
# Pantone 659C - Medium Blue
PANTONE_659C = "#0EA5E9"  # Medium blue

# Dark theme color scheme dictionary
INOSYS_COLORS = {
    "light_blue": PANTONE_3597C,
    "dark_blue": PANTONE_7683C,
    "medium_blue": PANTONE_659C,
    "white": "#FFFFFF",
    "light_gray": "#2B2B2B",  # Dark gray for frames
    "dark_gray": "#1A1A1A",  # Very dark gray for background
    "text_primary": "#FFFFFF",  # White text for primary
    "text_secondary": "#E5E5E5",  # Light gray text for secondary
    "background_primary": "#1A1A1A",  # Main dark background
    "background_secondary": "#2B2B2B",  # Secondary dark background
    "background_tertiary": "#3A3A3A"  # Tertiary dark background
}


class ToothDiagram(ctk.CTkFrame):
    def __init__(self, parent: ctk.CTkFrame, callback: Callable[[List[int]], None]) -> None:
        super().__init__(parent)
        self.callback: Callable[[List[int]], None] = callback
        self.selected_teeth: List[int] = []
        self.tooth_buttons: Dict[int, ctk.CTkButton] = {}

        # Tooth numbering (Universal Numeric Notation)
        # Upper teeth: 1-16 (left to right)
        # Lower teeth: 32-17 (left to right)

        self.create_tooth_diagram()

    def create_tooth_diagram(self) -> None:
        title: ctk.CTkLabel = ctk.CTkLabel(
            self,
            text="Select Teeth (Universal Numeric Notation) - Click multiple teeth to select",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=INOSYS_COLORS["light_blue"]
        )
        title.pack(pady=10)

        # Clear selection button
        clear_button: ctk.CTkButton = ctk.CTkButton(
            self,
            text="Clear Selection",
            command=self.clear_selection,
            height=30,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=INOSYS_COLORS["dark_blue"],
            hover_color=INOSYS_COLORS["medium_blue"],
            text_color=INOSYS_COLORS["white"]
        )
        clear_button.pack(pady=(0, 10))

        # Upper teeth frame
        upper_frame: ctk.CTkFrame = ctk.CTkFrame(self, fg_color=INOSYS_COLORS["background_tertiary"])
        upper_frame.pack(pady=5)

        upper_label: ctk.CTkLabel = ctk.CTkLabel(
            upper_frame,
            text="Upper Teeth",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=INOSYS_COLORS["text_primary"]
        )
        upper_label.pack(pady=5)

        upper_teeth_frame: ctk.CTkFrame = ctk.CTkFrame(upper_frame, fg_color=INOSYS_COLORS["background_secondary"])
        upper_teeth_frame.pack(pady=5)

        # Upper teeth (1-16, left to right)
        for i in range(1, 17):
            btn: ctk.CTkButton = ctk.CTkButton(
                upper_teeth_frame,
                text=str(i),
                width=40,
                height=40,
                command=lambda tooth=i: self.select_tooth(tooth),
                fg_color=INOSYS_COLORS["dark_blue"],
                hover_color=INOSYS_COLORS["medium_blue"],
                text_color=INOSYS_COLORS["white"],
                font=ctk.CTkFont(size=12, weight="bold"),
                border_width=1,
                border_color=INOSYS_COLORS["light_blue"]
            )
            btn.grid(row=0, column=i - 1, padx=2, pady=2)
            self.tooth_buttons[i] = btn

        # Lower teeth frame
        lower_frame: ctk.CTkFrame = ctk.CTkFrame(self, fg_color=INOSYS_COLORS["background_tertiary"])
        lower_frame.pack(pady=5)

        lower_label: ctk.CTkLabel = ctk.CTkLabel(
            lower_frame,
            text="Lower Teeth",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=INOSYS_COLORS["text_primary"]
        )
        lower_label.pack(pady=5)

        lower_teeth_frame: ctk.CTkFrame = ctk.CTkFrame(lower_frame, fg_color=INOSYS_COLORS["background_secondary"])
        lower_teeth_frame.pack(pady=5)

        # Lower teeth (32-17, left to right)
        for i in range(32, 16, -1):
            btn: ctk.CTkButton = ctk.CTkButton(
                lower_teeth_frame,
                text=str(i),
                width=40,
                height=40,
                command=lambda tooth=i: self.select_tooth(tooth),
                fg_color=INOSYS_COLORS["dark_blue"],
                hover_color=INOSYS_COLORS["medium_blue"],
                text_color=INOSYS_COLORS["white"],
                font=ctk.CTkFont(size=12, weight="bold"),
                border_width=1,
                border_color=INOSYS_COLORS["light_blue"]
            )
            btn.grid(row=0, column=32 - i, padx=2, pady=2)
            self.tooth_buttons[i] = btn

    def select_tooth(self, tooth_num: int) -> None:
        # Toggle tooth selection
        if tooth_num in self.selected_teeth:
            # Deselect tooth
            self.selected_teeth.remove(tooth_num)
            self.tooth_buttons[tooth_num].configure(
                fg_color=INOSYS_COLORS["dark_blue"],
                border_color=INOSYS_COLORS["light_blue"],
                border_width=1
            )
        else:
            # Select tooth
            self.selected_teeth.append(tooth_num)
            self.tooth_buttons[tooth_num].configure(
                fg_color=INOSYS_COLORS["light_blue"],
                border_color=INOSYS_COLORS["white"],
                border_width=2
            )

        # Sort the selected teeth list for consistent display
        self.selected_teeth.sort()
        self.callback(self.selected_teeth.copy())

    def clear_selection(self) -> None:
        # Reset all buttons to default color
        for btn in self.tooth_buttons.values():
            btn.configure(
                fg_color=INOSYS_COLORS["dark_blue"],
                border_color=INOSYS_COLORS["light_blue"],
                border_width=1
            )

        # Clear selection list
        self.selected_teeth.clear()
        self.callback(self.selected_teeth.copy())


class PrimusImplantApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        self.title("Primus Dental Implant Report Generator")
        self.geometry("1200x900")

        # Set window icon
        self.set_window_icon()

        # Set taskbar icon (Windows-specific)
        self.set_taskbar_icon()

        # Configure window colors
        self.configure(fg_color=INOSYS_COLORS["background_primary"])

        # Initialize instance variables with type hints
        self.implant_data: pd.DataFrame = pd.DataFrame()
        self.implant_plans: List[Dict[str, Any]] = []

        # GUI components - will be initialized in create_widgets
        self.notebook: ctk.CTkTabview
        self.tooth_diagram: ToothDiagram
        self.selected_teeth_label: ctk.CTkLabel
        self.implant_line_var: ctk.StringVar
        self.implant_line_combo: ctk.CTkComboBox
        self.implant_diameter_var: ctk.StringVar
        self.implant_diameter_combo: ctk.CTkComboBox
        self.implant_length_var: ctk.StringVar
        self.implant_length_combo: ctk.CTkComboBox
        self.offset_var: ctk.StringVar
        self.offset_combo: ctk.CTkComboBox
        self.surgical_approach_var: ctk.StringVar
        self.plan_scrollable_frame: ctk.CTkScrollableFrame
        self.doctor_name_entry: ctk.CTkEntry
        self.patient_name_entry: ctk.CTkEntry
        self.case_number_entry: ctk.CTkEntry

        # Load CSV data
        self.load_implant_data()

        self.create_widgets()
        self.create_menu()

    def load_implant_data(self) -> None:
        """Load implant data from CSV file"""
        csv_filename: str = "Primus Implant List - Primus Implant List.csv"

        try:
            if not os.path.exists(csv_filename):
                raise FileNotFoundError(f"CSV file '{csv_filename}' not found in current directory")

            self.implant_data = pd.read_csv(csv_filename)
            print(f"Implant data loaded successfully from {csv_filename}!")
            print(f"Total records: {len(self.implant_data)}")

            # Validate required columns
            required_columns: List[str] = [
                'Implant Line', 'Implant Part No', 'Implant Diameter', 'Implant Length',
                'Guide Sleeve', 'Drill Length', 'Offset', 'Starter Drill',
                'Initial Drill 1', 'Initial Drill 2', 'Drill 1', 'Drill 2', 'Drill 3', 'Drill 4'
            ]

            missing_columns: List[str] = [col for col in required_columns if col not in self.implant_data.columns]
            if missing_columns:
                raise ValueError(f"Missing required columns: {missing_columns}")

        except FileNotFoundError as e:
            messagebox.showerror("File Not Found", str(e))
            self.implant_data = pd.DataFrame()
        except pd.errors.EmptyDataError:
            messagebox.showerror("Error", f"The file '{csv_filename}' is empty")
            self.implant_data = pd.DataFrame()
        except pd.errors.ParserError as e:
            messagebox.showerror("Error", f"Failed to parse CSV file: {str(e)}")
            self.implant_data = pd.DataFrame()
        except ValueError as e:
            messagebox.showerror("Data Validation Error", str(e))
            self.implant_data = pd.DataFrame()
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error loading implant data: {str(e)}")
            self.implant_data = pd.DataFrame()

    def create_widgets(self) -> None:
        # Main container
        main_frame: ctk.CTkFrame = ctk.CTkFrame(self, fg_color=INOSYS_COLORS["background_secondary"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Header frame for logo and title
        header_frame: ctk.CTkFrame = ctk.CTkFrame(main_frame, fg_color=INOSYS_COLORS["medium_blue"])
        header_frame.pack(fill="x", pady=(0, 20))

        # Load and display logo
        self.load_and_display_logo(header_frame)

        # Title - vertically centered
        title_label: ctk.CTkLabel = ctk.CTkLabel(
            header_frame,
            text="Primus Dental Implant Report Generator",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=INOSYS_COLORS["white"]
        )
        title_label.pack(expand=True, pady=10)

        # Create notebook for tabs
        self.notebook = ctk.CTkTabview(
            main_frame,
            fg_color=INOSYS_COLORS["background_secondary"],
            segmented_button_fg_color=INOSYS_COLORS["background_tertiary"],
            segmented_button_selected_color=INOSYS_COLORS["light_blue"],
            segmented_button_selected_hover_color=INOSYS_COLORS["medium_blue"],
            text_color=INOSYS_COLORS["text_primary"],
            segmented_button_unselected_color=INOSYS_COLORS["background_tertiary"],
            segmented_button_unselected_hover_color=INOSYS_COLORS["background_secondary"]
        )
        self.notebook.pack(fill="both", expand=True, padx=20, pady=10)

        # Add tabs
        self.notebook.add("Add Implant")
        self.notebook.add("Review Plan")
        self.notebook.add("Generate Report")

        # Setup each tab
        self.setup_add_implant_tab()
        self.setup_review_plan_tab()
        self.setup_generate_report_tab()

    def load_and_display_logo(self, parent_frame: ctk.CTkFrame) -> None:
        """Load and display the Inosys logo in the GUI"""
        logo_files: List[str] = [
            "inosys_logo.png",
            "inosys_logo.jpg",
            "inosys_logo.jpeg",
            "logo.png",
            "logo.jpg",
            "logo.jpeg"
        ]

        logo_path: Optional[str] = None
        for logo_file in logo_files:
            if os.path.exists(logo_file):
                logo_path = logo_file
                break

        if logo_path:
            try:
                # Load logo using CTkImage
                pil_image = PILImage.open(logo_path)
                # Resize logo to fit nicely in header (maintain aspect ratio)
                original_width, original_height = pil_image.size
                max_height = 80
                aspect_ratio = original_width / original_height
                new_width = int(max_height * aspect_ratio)

                logo_image = ctk.CTkImage(
                    light_image=pil_image,
                    dark_image=pil_image,
                    size=(new_width, max_height)
                )

                logo_label: ctk.CTkLabel = ctk.CTkLabel(
                    parent_frame,
                    image=logo_image,
                    text=""
                )
                logo_label.pack(side="left", padx=(10, 0), pady=10)

                print(f"Logo loaded successfully from {logo_path}")

            except Exception as e:
                print(f"Error loading logo from {logo_path}: {str(e)}")
        else:
            print("Logo file not found. Please ensure the logo is saved as one of: " + ", ".join(logo_files))

    def set_window_icon(self) -> None:
        """Set the window icon"""
        icon_files: List[str] = [
            "icon.ico",
            "icon.png",
            "window_icon.ico",
            "window_icon.png",
            "inosys_icon.ico",
            "inosys_icon.png",
            "logo.ico",
            "logo.png"
        ]

        for icon_file in icon_files:
            if os.path.exists(icon_file):
                try:
                    print(f"Found icon file: {icon_file}")

                    # Try ICO files first (most reliable for Windows)
                    if icon_file.lower().endswith('.ico'):
                        try:
                            self.iconbitmap(icon_file)
                            # Also try wm_iconbitmap for better compatibility
                            self.wm_iconbitmap(icon_file)
                            print(f"Window icon set successfully from {icon_file} (ICO)")
                            return
                        except Exception as e:
                            print(f"ICO method failed: {str(e)}")
                            continue

                    # For PNG files, try multiple methods
                    elif icon_file.lower().endswith('.png'):
                        success = False

                        # Method 1: Convert PNG to ICO and use iconbitmap
                        try:
                            # Convert PNG to ICO in memory
                            pil_image = PILImage.open(icon_file)
                            # Create multiple sizes for ICO
                            icon_sizes = [(16, 16), (32, 32), (48, 48)]
                            ico_path = icon_file.replace('.png', '_temp.ico')

                            # Save as temporary ICO file
                            pil_image.save(ico_path, format='ICO', sizes=icon_sizes)

                            # Use the ICO file
                            self.iconbitmap(ico_path)
                            self.wm_iconbitmap(ico_path)

                            # Clean up temporary file
                            try:
                                os.remove(ico_path)
                            except:
                                pass

                            print(f"Window icon set successfully from {icon_file} (PNG->ICO conversion)")
                            return

                        except Exception as e:
                            print(f"PNG->ICO conversion failed: {str(e)}")

                        # Method 2: Use PhotoImage with iconphoto
                        try:
                            from PIL import ImageTk

                            # Load and resize icon
                            pil_image = PILImage.open(icon_file)
                            # Try multiple sizes
                            for size in [(64, 64), (32, 32), (16, 16)]:
                                resized_image = pil_image.resize(size, PILImage.Resampling.LANCZOS)
                                photo = ImageTk.PhotoImage(resized_image)

                                # Try both iconphoto methods
                                self.iconphoto(True, photo)
                                self.wm_iconphoto(True, photo)

                                # Keep reference
                                if not hasattr(self, '_icon_photos'):
                                    self._icon_photos = []
                                self._icon_photos.append(photo)

                            print(f"Window icon set successfully from {icon_file} (PhotoImage)")
                            return

                        except Exception as e:
                            print(f"PhotoImage method failed: {str(e)}")

                        # Method 3: Simple tkinter PhotoImage
                        try:
                            import tkinter as tk
                            photo = tk.PhotoImage(file=icon_file)
                            self.iconphoto(True, photo)
                            self.wm_iconphoto(True, photo)
                            self._icon_photo = photo
                            print(f"Window icon set successfully from {icon_file} (simple PhotoImage)")
                            return
                        except Exception as e:
                            print(f"Simple PhotoImage method failed: {str(e)}")

                except Exception as e:
                    print(f"Error setting window icon from {icon_file}: {str(e)}")
                    continue

        print("Window icon could not be set. For best results, convert your icon to ICO format.")
        print("You can use online converters or tools like GIMP to convert PNG to ICO.")

    def set_taskbar_icon(self) -> None:
        """Set the taskbar icon (Windows-specific solution)"""
        try:
            # Try to set Windows taskbar icon using ctypes
            if sys.platform.startswith('win'):
                try:
                    import ctypes
                    from ctypes import wintypes

                    # Get the window handle
                    hwnd = self.winfo_id()

                    # Find icon file
                    icon_files = ["icon.ico", "icon.png"]
                    icon_path = None

                    for icon_file in icon_files:
                        if os.path.exists(icon_file):
                            icon_path = os.path.abspath(icon_file)
                            break

                    if icon_path and icon_path.endswith('.ico'):
                        # Load the icon
                        IMAGE_ICON = 1
                        LR_LOADFROMFILE = 0x00000010
                        LR_DEFAULTSIZE = 0x00000040

                        hicon = ctypes.windll.user32.LoadImageW(
                            None, icon_path, IMAGE_ICON, 0, 0,
                            LR_LOADFROMFILE | LR_DEFAULTSIZE
                        )

                        if hicon:
                            # Set both small and large icons
                            WM_SETICON = 0x0080
                            ICON_SMALL = 0
                            ICON_BIG = 1

                            ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon)
                            ctypes.windll.user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, hicon)

                            print("Taskbar icon set successfully using Windows API")
                            return

                except Exception as e:
                    print(f"Windows API method failed: {str(e)}")

                # Alternative method: Set application ID
                try:
                    import ctypes
                    app_id = "Inosys.PrimusDentalImplant.ReportGenerator.1.0"
                    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
                    print("Application ID set for taskbar grouping")
                except Exception as e:
                    print(f"Could not set application ID: {str(e)}")

        except Exception as e:
            print(f"Taskbar icon setup failed: {str(e)}")

        print("Note: Taskbar icons work best when the application is run as an executable (.exe)")
        print("Consider using PyInstaller with --icon=icon.ico for production deployment")

    def create_menu(self) -> None:
        """Create the application menu bar"""
        # Create menu bar
        menubar = tk.Menu(self)
        self.config(menu=menubar)

        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about_dialog)
        help_menu.add_separator()
        help_menu.add_command(label="Check for Updates", command=self.check_for_updates)

    def show_about_dialog(self) -> None:
        """Show the About dialog"""
        about_window = ctk.CTkToplevel(self)
        about_window.title("About")
        about_window.geometry("400x300")
        about_window.resizable(False, False)
        about_window.configure(fg_color=INOSYS_COLORS["background_primary"])

        # Center the window
        about_window.transient(self)
        about_window.grab_set()

        # Main frame
        main_frame = ctk.CTkFrame(about_window, fg_color=INOSYS_COLORS["background_secondary"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Logo (if available)
        self.add_logo_to_about(main_frame)

        # Application name
        app_name_label = ctk.CTkLabel(
            main_frame,
            text="Primus Dental Implant\nReport Generator",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=INOSYS_COLORS["light_blue"]
        )
        app_name_label.pack(pady=(10, 5))

        # Version information
        version_label = ctk.CTkLabel(
            main_frame,
            text=f"Version {APP_VERSION}",
            font=ctk.CTkFont(size=12),
            text_color=INOSYS_COLORS["text_primary"]
        )
        version_label.pack(pady=2)

        build_label = ctk.CTkLabel(
            main_frame,
            text=f"Build Date: {APP_BUILD_DATE}",
            font=ctk.CTkFont(size=10),
            text_color=INOSYS_COLORS["text_secondary"]
        )
        build_label.pack(pady=2)

        # Company information
        company_label = ctk.CTkLabel(
            main_frame,
            text="© 2025 Inosys Implant\nAll rights reserved.",
            font=ctk.CTkFont(size=10),
            text_color=INOSYS_COLORS["text_secondary"]
        )
        company_label.pack(pady=(15, 10))

        # System information
        python_version = f"Python {sys.version.split()[0]}"
        system_label = ctk.CTkLabel(
            main_frame,
            text=f"Runtime: {python_version}",
            font=ctk.CTkFont(size=9),
            text_color=INOSYS_COLORS["text_secondary"]
        )
        system_label.pack(pady=2)

        # Close button
        close_button = ctk.CTkButton(
            main_frame,
            text="Close",
            command=about_window.destroy,
            width=100,
            height=30,
            fg_color=INOSYS_COLORS["medium_blue"],
            hover_color=INOSYS_COLORS["light_blue"]
        )
        close_button.pack(pady=(15, 10))

    def add_logo_to_about(self, parent_frame: ctk.CTkFrame) -> None:
        """Add logo to about dialog if available"""
        logo_files = ["inosys_logo.png", "logo.png", "icon.png"]

        for logo_file in logo_files:
            if os.path.exists(logo_file):
                try:
                    pil_image = PILImage.open(logo_file)
                    # Resize for about dialog
                    original_width, original_height = pil_image.size
                    aspect_ratio = original_width / original_height
                    new_height = 60
                    new_width = int(new_height * aspect_ratio)

                    logo_image = ctk.CTkImage(
                        light_image=pil_image,
                        dark_image=pil_image,
                        size=(new_width, new_height)
                    )

                    logo_label = ctk.CTkLabel(parent_frame, image=logo_image, text="")
                    logo_label.pack(pady=(10, 0))
                    return
                except Exception as e:
                    print(f"Could not load logo for about dialog: {e}")
                    continue

    def check_for_updates(self) -> None:
        """Check for updates and initiate download if available"""
        # Show checking dialog
        checking_dialog = self.show_update_dialog("Checking for updates...", show_progress=True)

        # Run update check in separate thread to avoid blocking UI
        update_thread = threading.Thread(target=self._update_check_worker, args=(checking_dialog,))
        update_thread.daemon = True
        update_thread.start()

    def _update_check_worker(self, checking_dialog: ctk.CTkToplevel) -> None:
        """Worker thread for checking and downloading updates"""
        try:
            update_path = os.path.join(UPDATE_SERVER_PATH, UPDATE_FILENAME)

            # Check if update file exists
            if not os.path.exists(update_path):
                self.after(100, lambda: self._show_update_result(checking_dialog, "no_update"))
                return

            # Get current executable path
            if getattr(sys, 'frozen', False):
                current_exe = sys.executable
            else:
                # For development, use a placeholder
                current_exe = "main.py"

            # Compare file sizes or modification times (simple update detection)
            try:
                update_stat = os.stat(update_path)
                current_stat = os.stat(current_exe) if os.path.exists(current_exe) else None

                if current_stat and update_stat.st_mtime <= current_stat.st_mtime:
                    self.after(100, lambda: self._show_update_result(checking_dialog, "up_to_date"))
                    return
            except:
                pass  # If we can't compare, proceed with update

            # Update available, start download
            self.after(100, lambda: self._show_update_result(checking_dialog, "update_available", update_path))

        except Exception as e:
            error_msg = f"Update check failed: {str(e)}"
            self.after(100, lambda: self._show_update_result(checking_dialog, "error", error_msg))

    def _show_update_result(self, checking_dialog: ctk.CTkToplevel, result: str, data: Optional[str] = None) -> None:
        """Show the result of update check"""
        checking_dialog.destroy()

        if result == "no_update":
            messagebox.showinfo("No Updates", "No updates are available at this time.")
        elif result == "up_to_date":
            messagebox.showinfo("Up to Date", "You are running the latest version.")
        elif result == "error":
            messagebox.showerror("Update Error", data or "An error occurred while checking for updates.")
        elif result == "update_available":
            if messagebox.askyesno("Update Available",
                                   "A new version is available. Would you like to download and install it now?\n\n"
                                   "The application will restart after the update."):
                self._download_and_install_update(data)

    def _download_and_install_update(self, update_path: str) -> None:
        """Download and install the update"""
        progress_dialog = self.show_update_dialog("Downloading update...", show_progress=True)

        # Run download in separate thread
        download_thread = threading.Thread(target=self._download_worker, args=(update_path, progress_dialog))
        download_thread.daemon = True
        download_thread.start()

    def _download_worker(self, update_path: str, progress_dialog: ctk.CTkToplevel) -> None:
        """Worker thread for downloading and installing update"""
        try:
            # Get current executable info
            if getattr(sys, 'frozen', False):
                current_exe = sys.executable
                current_dir = os.path.dirname(current_exe)
                current_name = os.path.basename(current_exe)
            else:
                # For development mode
                self.after(100, lambda: self._show_download_result(progress_dialog, "error",
                                                                   "Updates are only available for compiled executables."))
                return

            # Create backup of current executable
            backup_path = os.path.join(current_dir, f"{current_name}.backup")
            shutil.copy2(current_exe, backup_path)

            # Copy new executable
            temp_path = os.path.join(current_dir, f"{current_name}.new")
            shutil.copy2(update_path, temp_path)

            # Create update script
            update_script = self._create_update_script(current_exe, temp_path, backup_path)

            self.after(100, lambda: self._show_download_result(progress_dialog, "success", update_script))

        except Exception as e:
            error_msg = f"Update installation failed: {str(e)}"
            self.after(100, lambda: self._show_download_result(progress_dialog, "error", error_msg))

    def _create_update_script(self, current_exe: str, temp_path: str, backup_path: str) -> str:
        """Create a batch script to complete the update"""
        script_content = f'''@echo off
echo Updating Primus Dental Implant Report Generator...
timeout /t 2 /nobreak >nul

REM Replace the executable
move "{temp_path}" "{current_exe}"

REM Update shortcuts
set "desktop_shortcut=%USERPROFILE%\\Desktop\\Primus Report Generator.lnk"
set "startmenu_shortcut=%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\Inosys\\Primus Report Generator.lnk"

REM Update desktop shortcut if it exists
if exist "%desktop_shortcut%" (
    del "%desktop_shortcut%"
    powershell "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%desktop_shortcut%'); $s.TargetPath = '{current_exe}'; $s.Save()"
)

REM Update start menu shortcut if it exists
if exist "%startmenu_shortcut%" (
    del "%startmenu_shortcut%"
    powershell "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%startmenu_shortcut%'); $s.TargetPath = '{current_exe}'; $s.Save()"
)

REM Clean up backup after successful update
if exist "{current_exe}" (
    del "{backup_path}"
) else (
    REM Restore backup if update failed
    move "{backup_path}" "{current_exe}"
)

REM Restart the application
start "" "{current_exe}"

REM Delete this script
del "%~f0"
'''

        script_path = os.path.join(os.path.dirname(current_exe), "update_primus.bat")
        with open(script_path, 'w') as f:
            f.write(script_content)

        return script_path

    def _show_download_result(self, progress_dialog: ctk.CTkToplevel, result: str, data: str) -> None:
        """Show the result of download/installation"""
        progress_dialog.destroy()

        if result == "error":
            messagebox.showerror("Update Error", data)
        elif result == "success":
            if messagebox.showinfo("Update Ready",
                                   "Update downloaded successfully. The application will now restart to complete the installation."):
                # Execute the update script and exit
                subprocess.Popen(data, shell=True)
                self.quit()

    def show_update_dialog(self, message: str, show_progress: bool = False) -> ctk.CTkToplevel:
        """Show a dialog for update operations"""
        dialog = ctk.CTkToplevel(self)
        dialog.title("Update")
        dialog.geometry("300x150")
        dialog.resizable(False, False)
        dialog.configure(fg_color=INOSYS_COLORS["background_primary"])

        # Center the dialog
        dialog.transient(self)
        dialog.grab_set()

        # Main frame
        main_frame = ctk.CTkFrame(dialog, fg_color=INOSYS_COLORS["background_secondary"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Message
        message_label = ctk.CTkLabel(
            main_frame,
            text=message,
            font=ctk.CTkFont(size=12),
            text_color=INOSYS_COLORS["text_primary"]
        )
        message_label.pack(pady=20)

        # Progress bar (if requested)
        if show_progress:
            progress = ctk.CTkProgressBar(main_frame, width=200)
            progress.pack(pady=10)
            progress.configure(mode="indeterminate")
            progress.start()

        return dialog

    def setup_add_implant_tab(self) -> None:
        tab: ctk.CTkFrame = self.notebook.tab("Add Implant")

        # Create scrollable frame
        scrollable_frame: ctk.CTkScrollableFrame = ctk.CTkScrollableFrame(
            tab,
            fg_color=INOSYS_COLORS["background_secondary"]
        )
        scrollable_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Tooth selection
        tooth_frame: ctk.CTkFrame = ctk.CTkFrame(scrollable_frame, fg_color=INOSYS_COLORS["background_tertiary"])
        tooth_frame.pack(fill="x", padx=10, pady=10)

        self.tooth_diagram = ToothDiagram(tooth_frame, self.on_teeth_selected)
        self.tooth_diagram.pack(fill="both", expand=True, padx=10, pady=10)

        # Selected teeth display
        self.selected_teeth_label = ctk.CTkLabel(
            scrollable_frame,
            text="Selected Teeth: None",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=INOSYS_COLORS["light_blue"]
        )
        self.selected_teeth_label.pack(pady=10)

        # Input fields frame
        input_frame: ctk.CTkFrame = ctk.CTkFrame(scrollable_frame, fg_color=INOSYS_COLORS["background_tertiary"])
        input_frame.pack(fill="x", padx=10, pady=10)

        # Implant Line
        ctk.CTkLabel(
            input_frame,
            text="Implant Line:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=INOSYS_COLORS["text_primary"]
        ).grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.implant_line_var = ctk.StringVar(value="Primus")
        self.implant_line_combo = ctk.CTkComboBox(
            input_frame,
            values=["Primus"],
            variable=self.implant_line_var,
            state="readonly",
            fg_color=INOSYS_COLORS["background_secondary"],
            button_color=INOSYS_COLORS["medium_blue"],
            button_hover_color=INOSYS_COLORS["light_blue"],
            text_color=INOSYS_COLORS["text_primary"]
        )
        self.implant_line_combo.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        # Implant Diameter
        ctk.CTkLabel(
            input_frame,
            text="Implant Diameter:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=INOSYS_COLORS["text_primary"]
        ).grid(row=1, column=0, padx=10, pady=5, sticky="w")

        self.implant_diameter_var = ctk.StringVar()
        self.implant_diameter_combo = ctk.CTkComboBox(
            input_frame,
            values=["3.5", "4.0", "4.5", "5.0"],
            variable=self.implant_diameter_var,
            fg_color=INOSYS_COLORS["background_secondary"],
            button_color=INOSYS_COLORS["medium_blue"],
            button_hover_color=INOSYS_COLORS["light_blue"],
            text_color=INOSYS_COLORS["text_primary"]
        )
        self.implant_diameter_combo.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        # Implant Length
        ctk.CTkLabel(
            input_frame,
            text="Implant Length:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=INOSYS_COLORS["text_primary"]
        ).grid(row=2, column=0, padx=10, pady=5, sticky="w")

        self.implant_length_var = ctk.StringVar()
        self.implant_length_combo = ctk.CTkComboBox(
            input_frame,
            values=["7.5", "8.5", "10.0", "11.5", "13.0"],
            variable=self.implant_length_var,
            fg_color=INOSYS_COLORS["background_secondary"],
            button_color=INOSYS_COLORS["medium_blue"],
            button_hover_color=INOSYS_COLORS["light_blue"],
            text_color=INOSYS_COLORS["text_primary"]
        )
        self.implant_length_combo.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # Offset
        ctk.CTkLabel(
            input_frame,
            text="Offset:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=INOSYS_COLORS["text_primary"]
        ).grid(row=3, column=0, padx=10, pady=5, sticky="w")

        self.offset_var = ctk.StringVar()
        self.offset_combo = ctk.CTkComboBox(
            input_frame,
            values=["10", "11.5", "13"],
            variable=self.offset_var,
            fg_color=INOSYS_COLORS["background_secondary"],
            button_color=INOSYS_COLORS["medium_blue"],
            button_hover_color=INOSYS_COLORS["light_blue"],
            text_color=INOSYS_COLORS["text_primary"]
        )
        self.offset_combo.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        # Surgical Approach (Flap/Flapless)
        ctk.CTkLabel(
            input_frame,
            text="Surgical Approach:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=INOSYS_COLORS["text_primary"]
        ).grid(row=4, column=0, padx=10, pady=5, sticky="w")

        # Radio button frame
        approach_frame = ctk.CTkFrame(input_frame, fg_color="transparent")
        approach_frame.grid(row=4, column=1, padx=10, pady=5, sticky="ew")

        self.surgical_approach_var = ctk.StringVar(value="flapless")

        flap_radio = ctk.CTkRadioButton(
            approach_frame,
            text="Flap",
            variable=self.surgical_approach_var,
            value="flap",
            text_color=INOSYS_COLORS["text_primary"],
            fg_color=INOSYS_COLORS["medium_blue"],
            hover_color=INOSYS_COLORS["light_blue"]
        )
        flap_radio.pack(side="left", padx=(0, 20))

        flapless_radio = ctk.CTkRadioButton(
            approach_frame,
            text="Flapless",
            variable=self.surgical_approach_var,
            value="flapless",
            text_color=INOSYS_COLORS["text_primary"],
            fg_color=INOSYS_COLORS["medium_blue"],
            hover_color=INOSYS_COLORS["light_blue"]
        )
        flapless_radio.pack(side="left")

        # Configure grid weights
        input_frame.columnconfigure(1, weight=1)

        # Add implant button
        add_button: ctk.CTkButton = ctk.CTkButton(
            scrollable_frame,
            text="Add Implants to Plan",
            command=self.add_implants_to_plan,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=INOSYS_COLORS["light_blue"],
            hover_color=INOSYS_COLORS["medium_blue"],
            text_color=INOSYS_COLORS["white"]
        )
        add_button.pack(pady=20)

    def setup_review_plan_tab(self) -> None:
        tab: ctk.CTkFrame = self.notebook.tab("Review Plan")

        # Frame for the plan list
        plan_frame: ctk.CTkFrame = ctk.CTkFrame(tab, fg_color=INOSYS_COLORS["background_secondary"])
        plan_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Title
        title_label: ctk.CTkLabel = ctk.CTkLabel(
            plan_frame,
            text="Current Implant Plan",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=INOSYS_COLORS["light_blue"]
        )
        title_label.pack(pady=10)

        # Scrollable frame for plan items
        self.plan_scrollable_frame = ctk.CTkScrollableFrame(
            plan_frame,
            fg_color=INOSYS_COLORS["background_tertiary"]
        )
        self.plan_scrollable_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Clear plan button
        clear_button: ctk.CTkButton = ctk.CTkButton(
            plan_frame,
            text="Clear All Plans",
            command=self.clear_all_plans,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=INOSYS_COLORS["dark_blue"],
            hover_color=INOSYS_COLORS["medium_blue"],
            text_color=INOSYS_COLORS["white"]
        )
        clear_button.pack(pady=10)

    def setup_generate_report_tab(self) -> None:
        tab: ctk.CTkFrame = self.notebook.tab("Generate Report")

        # Frame for report generation
        report_frame: ctk.CTkFrame = ctk.CTkFrame(tab, fg_color=INOSYS_COLORS["background_secondary"])
        report_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Title
        title_label: ctk.CTkLabel = ctk.CTkLabel(
            report_frame,
            text="Generate PDF Report",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=INOSYS_COLORS["light_blue"]
        )
        title_label.pack(pady=20)

        # Doctor information
        info_frame: ctk.CTkFrame = ctk.CTkFrame(report_frame, fg_color=INOSYS_COLORS["background_tertiary"])
        info_frame.pack(fill="x", padx=20, pady=10)

        ctk.CTkLabel(
            info_frame,
            text="Doctor Name:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=INOSYS_COLORS["text_primary"]
        ).grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.doctor_name_entry = ctk.CTkEntry(
            info_frame,
            width=300,
            fg_color=INOSYS_COLORS["background_secondary"],
            text_color=INOSYS_COLORS["text_primary"],
            border_color=INOSYS_COLORS["medium_blue"]
        )
        self.doctor_name_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(
            info_frame,
            text="Patient Name:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=INOSYS_COLORS["text_primary"]
        ).grid(row=1, column=0, padx=10, pady=5, sticky="w")

        self.patient_name_entry = ctk.CTkEntry(
            info_frame,
            width=300,
            fg_color=INOSYS_COLORS["background_secondary"],
            text_color=INOSYS_COLORS["text_primary"],
            border_color=INOSYS_COLORS["medium_blue"]
        )
        self.patient_name_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(
            info_frame,
            text="Case Number:",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=INOSYS_COLORS["text_primary"]
        ).grid(row=2, column=0, padx=10, pady=5, sticky="w")

        self.case_number_entry = ctk.CTkEntry(
            info_frame,
            width=300,
            fg_color=INOSYS_COLORS["background_secondary"],
            text_color=INOSYS_COLORS["text_primary"],
            border_color=INOSYS_COLORS["medium_blue"]
        )
        self.case_number_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        info_frame.columnconfigure(1, weight=1)

        # Generate button
        generate_button: ctk.CTkButton = ctk.CTkButton(
            report_frame,
            text="Generate PDF Report",
            command=self.generate_pdf_report,
            height=50,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=INOSYS_COLORS["light_blue"],
            hover_color=INOSYS_COLORS["medium_blue"],
            text_color=INOSYS_COLORS["white"]
        )
        generate_button.pack(pady=30)

    def on_teeth_selected(self, selected_teeth: List[int]) -> None:
        if selected_teeth:
            teeth_text = ", ".join(map(str, selected_teeth))
            self.selected_teeth_label.configure(text=f"Selected Teeth: {teeth_text}")
        else:
            self.selected_teeth_label.configure(text="Selected Teeth: None")

    def add_implants_to_plan(self) -> None:
        # Validate inputs
        if not self.tooth_diagram.selected_teeth:
            messagebox.showerror("Error", "Please select at least one tooth first!")
            return

        if not all([self.implant_diameter_var.get(), self.implant_length_var.get(), self.offset_var.get(),
                    self.surgical_approach_var.get()]):
            messagebox.showerror("Error", "Please fill in all fields!")
            return

        # Check if implant configuration exists in database
        diameter: float = float(self.implant_diameter_var.get())
        length: float = float(self.implant_length_var.get())
        offset: float = float(self.offset_var.get())

        matching_implant: pd.DataFrame = self.implant_data[
            (self.implant_data['Implant Diameter'] == diameter) &
            (self.implant_data['Implant Length'] == length) &
            (self.implant_data['Offset'] == offset)
            ]

        if matching_implant.empty:
            messagebox.showerror("Error", "No matching implant found in database for the selected specifications!")
            return

        # Check if the implant/drill length combination is valid
        implant_row = matching_implant.iloc[0]
        drill_fields = ['Starter Drill', 'Initial Drill 1', 'Initial Drill 2', 'Drill 1', 'Drill 2', 'Drill 3',
                        'Drill 4']

        # Check if any drill field contains 'x' (indicating invalid combination)
        invalid_drills = []
        for field in drill_fields:
            if str(implant_row[field]).lower().strip() == 'x':
                invalid_drills.append(field)

        if invalid_drills:
            invalid_fields_text = ", ".join(invalid_drills)
            messagebox.showerror(
                "Invalid Implant/Drill Length Combination",
                f"The selected implant length ({length}mm) and drill length ({implant_row['Drill Length']}mm) "
                f"combination is not compatible.\n\n"
                f"Invalid drill stages: {invalid_fields_text}\n\n"
                f"Please select a different implant length or offset to find a valid drilling protocol."
            )
            return

        # Track which teeth were added and which were replaced
        added_teeth: List[int] = []
        replaced_teeth: List[int] = []

        # Create implant plans for each selected tooth
        for tooth_number in self.tooth_diagram.selected_teeth:
            implant_plan: Dict[str, Any] = {
                'tooth_number': tooth_number,
                'implant_line': self.implant_line_var.get(),
                'diameter': diameter,
                'length': length,
                'offset': offset,
                'surgical_approach': self.surgical_approach_var.get(),
                'implant_data': implant_row.to_dict()
            }

            # Check if tooth already has an implant planned
            existing_plan: Optional[Dict[str, Any]] = next(
                (plan for plan in self.implant_plans if plan['tooth_number'] == tooth_number),
                None
            )

            if existing_plan:
                self.implant_plans.remove(existing_plan)
                replaced_teeth.append(tooth_number)
            else:
                added_teeth.append(tooth_number)

            self.implant_plans.append(implant_plan)

        # Update display and show success message
        self.update_plan_display()

        # Create success message
        message_parts: List[str] = []
        if added_teeth:
            message_parts.append(f"Added implants for teeth: {', '.join(map(str, added_teeth))}")
        if replaced_teeth:
            message_parts.append(f"Replaced existing implants for teeth: {', '.join(map(str, replaced_teeth))}")

        messagebox.showinfo("Success", "\n".join(message_parts))

        # Clear selections
        self.implant_diameter_var.set("")
        self.implant_length_var.set("")
        self.offset_var.set("")
        self.surgical_approach_var.set("flapless")
        self.tooth_diagram.clear_selection()

    def update_plan_display(self) -> None:
        # Clear existing plan display
        for widget in self.plan_scrollable_frame.winfo_children():
            widget.destroy()

        # Add each implant plan to display
        for i, plan in enumerate(self.implant_plans):
            plan_frame: ctk.CTkFrame = ctk.CTkFrame(
                self.plan_scrollable_frame,
                fg_color=INOSYS_COLORS["background_secondary"],
                border_width=1,
                border_color=INOSYS_COLORS["medium_blue"]
            )
            plan_frame.pack(fill="x", padx=10, pady=5)

            # Plan details
            approach_text = "Flap" if plan.get('surgical_approach', 'flap') == 'flap' else "Flapless"
            details_text: str = (
                f"Tooth {plan['tooth_number']}: {plan['implant_line']} - "
                f"{plan['diameter']}mm × {plan['length']}mm (Offset: {plan['offset']}mm) - {approach_text}"
            )
            details_label: ctk.CTkLabel = ctk.CTkLabel(
                plan_frame,
                text=details_text,
                font=ctk.CTkFont(size=12, weight="bold"),
                text_color=INOSYS_COLORS["text_primary"]
            )
            details_label.pack(side="left", padx=10, pady=10)

            # Remove button
            remove_button: ctk.CTkButton = ctk.CTkButton(
                plan_frame,
                text="Remove",
                width=80,
                command=lambda idx=i: self.remove_implant_plan(idx),
                fg_color=INOSYS_COLORS["dark_blue"],
                hover_color=INOSYS_COLORS["light_blue"],
                text_color=INOSYS_COLORS["white"]
            )
            remove_button.pack(side="right", padx=10, pady=10)

    def remove_implant_plan(self, index: int) -> None:
        if 0 <= index < len(self.implant_plans):
            self.implant_plans.pop(index)
            self.update_plan_display()

    def clear_all_plans(self) -> None:
        if messagebox.askyesno("Clear All Plans", "Are you sure you want to clear all implant plans?"):
            self.implant_plans.clear()
            self.update_plan_display()

    def generate_pdf_report(self) -> None:
        if not self.implant_plans:
            messagebox.showerror("Error", "No implant plans to generate report!")
            return

        # Get report information
        doctor_name: str = self.doctor_name_entry.get() or "Dr. [Name]"
        patient_name: str = self.patient_name_entry.get() or "[Patient Name]"
        case_number: str = self.case_number_entry.get() or "[Case Number]"

        # Ask for save location
        filename: str = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save Report As",
            initialfile=f"Primus_Report_{case_number}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        )

        if not filename:
            return

        try:
            self.create_pdf_report(filename, doctor_name, patient_name, case_number)
            messagebox.showinfo("Success", f"Report generated successfully!\nSaved as: {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate report: {str(e)}")

    def create_pdf_report(self, filename: str, doctor_name: str, patient_name: str, case_number: str) -> None:
        doc: SimpleDocTemplate = SimpleDocTemplate(
            filename,
            pagesize=letter,
            topMargin=0.5 * inch,
            bottomMargin=0.5 * inch,
            leftMargin=0.5 * inch,
            rightMargin=0.5 * inch,
            title="Drilling Protocol"
        )
        styles = getSampleStyleSheet()
        story: List[Any] = []

        # Custom styles
        title_style: ParagraphStyle = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=18,
            textColor=colors.Color(30 / 255, 58 / 255, 138 / 255),  # Pantone 7683C - Dark Blue
            alignment=TA_CENTER,
            spaceAfter=20
        )

        header_style: ParagraphStyle = ParagraphStyle(
            'CustomHeader',
            parent=styles['Heading2'],
            fontSize=12,
            textColor=colors.Color(30 / 255, 58 / 255, 138 / 255),  # Pantone 7683C - Dark Blue
            spaceBefore=10,
            spaceAfter=8
        )

        compact_style: ParagraphStyle = ParagraphStyle(
            'Compact',
            parent=styles['Normal'],
            fontSize=9,
            spaceBefore=2,
            spaceAfter=2
        )

        # Header section with logo and title
        header_data = []

        # Add logo to header if it exists
        logo_added: bool = self.add_logo_to_report_header(header_data)

        if header_data:
            # Create header table with logo and title
            header_table = Table(header_data, colWidths=[3 * inch, 4.5 * inch])
            header_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                ('ALIGN', (1, 0), (1, 0), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                ('TOPPADDING', (0, 0), (-1, -1), 0),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
            ]))
            story.append(header_table)
        else:
            # No logo, just title
            story.append(Paragraph("PRIMUS DENTAL IMPLANT SURGICAL DRILLING PROTOCOL", title_style))

        story.append(Spacer(1, 15))

        # Case information in compact format
        case_info: List[List[str]] = [
            ["Doctor:", doctor_name, "Date:", datetime.now().strftime("%B %d, %Y")],
            ["Patient:", patient_name, "Time:", datetime.now().strftime("%I:%M %p")],
            ["Case Number:", case_number, "Total Implants:", str(len(self.implant_plans))]
        ]

        case_table: Table = Table(case_info, colWidths=[1 * inch, 2.3 * inch, 1.2 * inch, 1.5 * inch])
        case_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTNAME', (2, 0), (2, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey),
            ('BACKGROUND', (0, 0), (0, -1), colors.Color(0 / 255, 181 / 255, 216 / 255)),  # Light blue
            ('BACKGROUND', (2, 0), (2, -1), colors.Color(0 / 255, 181 / 255, 216 / 255)),  # Light blue
        ]))

        story.append(case_table)
        story.append(Spacer(1, 15))

        # Horizontal line separator
        from reportlab.platypus import HRFlowable
        story.append(HRFlowable(width="100%", thickness=2, color=colors.Color(30 / 255, 58 / 255, 138 / 255)))
        story.append(Spacer(1, 10))

        # Sort implant plans by tooth number
        sorted_plans: List[Dict[str, Any]] = sorted(self.implant_plans, key=lambda x: x['tooth_number'])

        # Create comprehensive implant summary table
        story.append(Paragraph("IMPLANT SPECIFICATIONS & DRILLING PROTOCOL", header_style))

        # Main implant data table with all information
        implant_data = [
            ["Tooth", "Part No.", "Dia.", "Len.", "Offset", "Approach", "Guide", "Drill Len.", "Drilling Sequence"]]

        for i, plan in enumerate(sorted_plans):
            # Create drilling sequence string with line breaks for better fitting
            # Handle 'x' values in drilling sequence
            def format_drill_value(value):
                return "N/A" if str(value).lower().strip() == 'x' else f"{value}mm"

            drill_sequence = f"Start: {format_drill_value(plan['implant_data']['Starter Drill'])} → Init1: {format_drill_value(plan['implant_data']['Initial Drill 1'])} → Init2: {format_drill_value(plan['implant_data']['Initial Drill 2'])} →<br/>" \
                             f"D1: {format_drill_value(plan['implant_data']['Drill 1'])} → D2: {format_drill_value(plan['implant_data']['Drill 2'])} → D3: {format_drill_value(plan['implant_data']['Drill 3'])} → D4: {format_drill_value(plan['implant_data']['Drill 4'])}"

            approach_display = "Flap" if plan.get('surgical_approach', 'flap') == 'flap' else "Flapless"

            implant_data.append([
                str(plan['tooth_number']),
                plan['implant_data']['Implant Part No'],
                f"{plan['diameter']}mm",
                f"{plan['length']}mm",
                f"{plan['offset']}mm",
                approach_display,
                plan['implant_data']['Guide Sleeve'],
                f"{plan['implant_data']['Drill Length']}mm",
                Paragraph(drill_sequence, ParagraphStyle('DrillSeq',
                                                         parent=styles['Normal'],
                                                         fontSize=7,
                                                         leading=9))
            ])

        # Create table with appropriate column widths
        col_widths = [0.4 * inch, 1 * inch, 0.4 * inch, 0.4 * inch, 0.45 * inch, 0.55 * inch, 0.75 * inch, 0.6 * inch,
                      2.8 * inch]
        implant_table: Table = Table(implant_data, colWidths=col_widths, repeatRows=1)

        implant_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.Color(30 / 255, 58 / 255, 138 / 255)),  # Dark blue header
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))

        story.append(implant_table)
        story.append(Spacer(1, 15))

        # Horizontal line separator
        story.append(HRFlowable(width="100%", thickness=1, color=colors.Color(14 / 255, 165 / 255, 233 / 255)))
        story.append(Spacer(1, 10))

        # Compact surgical protocol in two columns
        story.append(Paragraph("SURGICAL PROTOCOL", header_style))

        protocol_data = [
            ["PRE-SURGICAL PREPARATION", "DRILLING PROTOCOL"],
            [
                "• Verify patient identity and surgical site\n"
                "• Confirm implant specifications\n"
                "• Prepare sterile surgical field\n"
                "• Check all instruments and drill bits\n"
                "• Ensure proper guide sleeve placement",

                "• Begin with the point drill in D4 bone\n"
                "• Use intermittent drilling (15-30 sec intervals)\n"
                "• Speed: 800-1200 RPM initial, 600-800 RPM final\n"
                "• Apply light pressure - let drill do the work\n"
                "• Use copious irrigation (minimum 50ml/min)\n"
                "• Check depth with gauge after each drill"
            ],
            ["POST-DRILLING VERIFICATION", "IMPORTANT NOTES"],
            [
                "• Irrigate osteotomy thoroughly\n"
                "• Check final depth and angulation\n"
                "• Verify diameter with sizing gauge\n"
                "• Ensure proper guide sleeve function\n"
                "• Proceed with implant placement protocol",

                "• Follow manufacturer drilling guidelines\n"
                "• Maintain sterile technique throughout\n"
                "• Account for offset measurements\n"
                "• Use depth gauge frequently\n"
                "• Document any deviations from protocol"
            ]
        ]

        protocol_table = Table(protocol_data, colWidths=[3.75 * inch, 3.75 * inch])
        protocol_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTNAME', (0, 2), (-1, 2), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey),
            ('BACKGROUND', (0, 0), (-1, 0), colors.Color(14 / 255, 165 / 255, 233 / 255)),  # Medium blue
            ('BACKGROUND', (0, 2), (-1, 2), colors.Color(14 / 255, 165 / 255, 233 / 255)),  # Medium blue
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('TEXTCOLOR', (0, 2), (-1, 2), colors.white),
        ]))

        story.append(protocol_table)
        story.append(Spacer(1, 15))

        # Final horizontal line
        story.append(HRFlowable(width="100%", thickness=2, color=colors.Color(30 / 255, 58 / 255, 138 / 255)))
        story.append(Spacer(1, 10))

        # Disclaimer section
        disclaimer_style: ParagraphStyle = ParagraphStyle(
            'Disclaimer',
            parent=styles['Normal'],
            fontSize=8,
            textColor=colors.Color(60 / 255, 60 / 255, 60 / 255),
            spaceBefore=10,
            spaceAfter=5,
            alignment=TA_LEFT,
            leading=10
        )

        story.append(Paragraph("<b>LIMITATION OF LIABILITY:</b>",
                               ParagraphStyle('DisclaimerTitle', parent=disclaimer_style, fontSize=9,
                                              textColor=colors.Color(30 / 255, 58 / 255, 138 / 255))))

        disclaimer_text = """This instruction incorporates a custom document that is based on a surgical plan proposed by the surgeon before operation. The surgeon, therefore, takes full medical responsibility for the design and the application of the surgical guide, the intended used surgical tray kit, implants and sleeves – all as specified on the order form received by the supplier. The custom document shall be considered as an addition to all other documents sent with and pertaining to the case, and it does not replace any of those other documents."""

        story.append(Paragraph(disclaimer_text, disclaimer_style))
        story.append(Spacer(1, 10))

        # Footer
        footer_text: str = f"""<b>Report Generated:</b> {datetime.now().strftime("%B %d, %Y at %I:%M %p")} | <b>Software:</b> Primus Implant Report Generator v1.0"""

        story.append(Paragraph(
            footer_text,
            ParagraphStyle('Footer', parent=styles['Normal'],
                           fontSize=7, textColor=colors.grey, alignment=TA_CENTER)
        ))

        # Build PDF
        doc.build(story)

    def add_logo_to_report_header(self, header_data: List[List[Any]]) -> bool:
        """Add logo to header data for table layout"""
        logo_files: List[str] = [
            "inosys_logo.png",
            "inosys_logo.jpg",
            "inosys_logo.jpeg",
            "logo.png",
            "logo.jpg",
            "logo.jpeg",
            "icon.png",
            "icon.jpg",
            "icon.jpeg"
        ]

        for logo_file in logo_files:
            if os.path.exists(logo_file):
                try:
                    # Get original image dimensions to maintain aspect ratio
                    with PILImage.open(logo_file) as pil_image:
                        original_width, original_height = pil_image.size
                        aspect_ratio = original_width / original_height

                        # Set desired height and calculate width to maintain aspect ratio
                        desired_height = 0.6 * inch
                        calculated_width = desired_height * aspect_ratio

                        # Create logo image for PDF with proper aspect ratio
                        logo_image = Image(logo_file, width=calculated_width, height=desired_height)

                        # Create title paragraph
                        title_para = Paragraph("PRIMUS DENTAL IMPLANT<br/><br/>SURGICAL DRILLING PROTOCOL",
                                               ParagraphStyle('HeaderTitle',
                                                              fontSize=16,
                                                              textColor=colors.Color(30 / 255, 58 / 255, 138 / 255),
                                                              alignment=TA_CENTER,
                                                              fontName='Helvetica-Bold',
                                                              leading=20))

                        header_data.append([logo_image, title_para])

                        print(f"Logo added to PDF header from {logo_file}")
                        return True

                except Exception as e:
                    print(f"Error adding logo to PDF header from {logo_file}: {str(e)}")
                    continue

        return False

    def add_logo_to_report(self, story: List[Any]) -> bool:
        """Add the Inosys logo to the PDF report"""
        logo_files: List[str] = [
            "inosys_logo.png",
            "inosys_logo.jpg",
            "inosys_logo.jpeg",
            "logo.png",
            "logo.jpg",
            "logo.jpeg",
            "icon.png",
            "icon.jpg",
            "icon.jpeg"
        ]

        for logo_file in logo_files:
            if os.path.exists(logo_file):
                try:
                    # Get original image dimensions to maintain aspect ratio
                    with PILImage.open(logo_file) as pil_image:
                        original_width, original_height = pil_image.size
                        aspect_ratio = original_width / original_height

                        # Set desired height and calculate width to maintain aspect ratio
                        desired_height = 0.8 * inch
                        calculated_width = desired_height * aspect_ratio

                        # Create logo image for PDF with proper aspect ratio
                        logo_image = Image(logo_file, width=calculated_width, height=desired_height)
                        logo_image.hAlign = 'LEFT'
                        story.append(logo_image)

                        print(f"Logo added to PDF report from {logo_file}")
                        print(f"Original dimensions: {original_width}x{original_height}")
                        print(f"PDF dimensions: {calculated_width / inch:.2f}\" x {desired_height / inch:.2f}\"")
                        print(f"Aspect ratio maintained: {aspect_ratio:.2f}")
                        return True

                except Exception as e:
                    print(f"Error adding logo to PDF from {logo_file}: {str(e)}")
                    continue

        print("Logo file not found for PDF report")
        return False


if __name__ == "__main__":
    app: PrimusImplantApp = PrimusImplantApp()
    app.mainloop()
