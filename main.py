import customtkinter as ctk
import pandas as pd
import tkinter as tk
from tkinter import messagebox, filedialog, scrolledtext
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import json
import os
import sys
import shutil
import subprocess
import threading
from datetime import datetime
from typing import Dict, List, Any, Optional, Callable, Tuple
from PIL import Image as PILImage
import tempfile
import webbrowser
from pathlib import Path

# Application version information
APP_VERSION = "1.0.4"  # Increment version
APP_BUILD_DATE = "2025-07-22"


# User-level paths
def get_user_app_directory() -> str:
    """Get user-specific application directory"""
    if sys.platform.startswith('win'):
        # Method 1: Use LOCALAPPDATA (most reliable)
        app_data = os.environ.get('LOCALAPPDATA')
        if app_data:
            return os.path.join(app_data, 'CreoDent', 'PrimusImplant')

        # Method 2: Use USERPROFILE (fallback)
        user_profile = os.environ.get('USERPROFILE')
        if user_profile:
            return os.path.join(user_profile, 'AppData', 'Local', 'CreoDent', 'PrimusImplant')

        # Method 3: Use expanduser (final fallback)
        return os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'CreoDent', 'PrimusImplant')
    else:
        return os.path.expanduser('~/.local/share/PrimusImplant')


def get_user_update_server_path() -> str:
    """Get user-level update server path"""
    # Primary: Network share for updates
    network_path = r"\\CDIMANQ30\Creoman-Active\CADCAM\Software\Primus Implant Report Generator\UserUpdates"

    # Fallback: Local shared folder
    local_path = r"C:\Shared\PrimusUpdates"

    # Check which path is accessible
    if os.path.exists(network_path):
        return network_path
    elif os.path.exists(local_path):
        return local_path
    else:
        # Final fallback: User directory updates subfolder
        return os.path.join(get_user_app_directory(), 'updates')


UPDATE_SERVER_PATH = get_user_update_server_path()
UPDATE_FILENAME = "Primus Implant Report Generator.exe"

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

        self.title("Primus Implant Report Generator")

        # Initialize logging first
        self.log_window_activity("=== APPLICATION STARTUP ===")
        self.log_window_activity(f"Python version: {sys.version}")
        self.log_window_activity(f"Platform: {sys.platform}")
        self.log_window_activity(f"App version: 1.0.4")

        # Load window geometry AFTER title is set
        self.load_window_geometry()

        # Set up user installation first
        self.setup_user_installation()

        # Set window icon
        self.set_window_icon()

        # Set taskbar icon (Windows-specific)
        self.set_taskbar_icon()

        # Configure window colors
        self.configure(fg_color=INOSYS_COLORS["background_primary"])

        # Initialize instance variables
        self.implant_data = pd.DataFrame()
        self.implant_plans = []
        self.current_case_notes = ""

        # Load CSV data
        self.load_implant_data()

        self.create_widgets()
        self.create_menu()

        # Bind window events
        self.protocol("WM_DELETE_WINDOW", self.on_window_close)

        self.log_window_activity("=== APPLICATION STARTUP COMPLETE ===")

        # Test window memory after a short delay
        self.after(1000, self.test_window_memory)

    def ensure_user_directories(self) -> None:
        """Ensure user directories exist"""
        try:
            user_dir = get_user_app_directory()
            os.makedirs(user_dir, exist_ok=True)

            # Create updates directory
            updates_dir = os.path.join(user_dir, 'updates')
            os.makedirs(updates_dir, exist_ok=True)

            # Create logs directory
            logs_dir = os.path.join(user_dir, 'logs')
            os.makedirs(logs_dir, exist_ok=True)

            print(f"User directories created at: {user_dir}")

        except Exception as e:
            print(f"Error creating user directories: {e}")

    def get_data_file_path(self, filename: str) -> str:
        """Get path to data file, checking user directory first"""
        user_dir = get_user_app_directory()
        user_file = os.path.join(user_dir, filename)

        # Check user directory first
        if os.path.exists(user_file):
            return user_file

        # Fall back to application directory
        if getattr(sys, 'frozen', False):
            app_dir = os.path.dirname(sys.executable)
        else:
            app_dir = os.path.dirname(os.path.abspath(__file__))

        return os.path.join(app_dir, filename)

    def copy_data_files_to_user_directory(self) -> None:
        """Copy data files to user directory if they don't exist"""
        try:
            user_dir = get_user_app_directory()

            if getattr(sys, 'frozen', False):
                app_dir = os.path.dirname(sys.executable)
            else:
                app_dir = os.path.dirname(os.path.abspath(__file__))

            data_files = [
                'Primus Implant List - Primus Implant List.csv',
                'inosys_logo.png',
                'icon.ico'
            ]

            for filename in data_files:
                src_path = os.path.join(app_dir, filename)
                dst_path = os.path.join(user_dir, filename)

                # Only copy if source exists and destination doesn't
                if os.path.exists(src_path) and not os.path.exists(dst_path):
                    shutil.copy2(src_path, dst_path)
                    print(f"Copied {filename} to user directory")

        except Exception as e:
            print(f"Error copying data files: {e}")

    def create_desktop_shortcut(self) -> None:
        """Create desktop shortcut for user installation"""
        try:
            if not sys.platform.startswith('win'):
                return

            if not getattr(sys, 'frozen', False):
                return  # Only for compiled versions

            current_exe = sys.executable
            desktop = os.path.join(os.path.expanduser('~'), 'Desktop')
            shortcut_path = os.path.join(desktop, "Primus Implant Report Generator.lnk")

            # Use PowerShell to create shortcut (no additional dependencies)
            ps_script = f'''
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("{shortcut_path}")
    $Shortcut.TargetPath = "{current_exe}"
    $Shortcut.WorkingDirectory = "{os.path.dirname(current_exe)}"
    $Shortcut.IconLocation = "{current_exe}"
    $Shortcut.Description = "Primus Implant Report Generator"
    $Shortcut.Save()
    '''

            result = subprocess.run(['powershell', '-Command', ps_script],
                                    capture_output=True, text=True)

            if result.returncode == 0:
                print("Desktop shortcut created successfully")
            else:
                print(f"Error creating shortcut: {result.stderr}")

        except Exception as e:
            print(f"Error creating desktop shortcut: {e}")

    def create_start_menu_shortcut(self) -> None:
        """Create Start Menu shortcut"""
        try:
            if not sys.platform.startswith('win'):
                return

            if not getattr(sys, 'frozen', False):
                return

            current_exe = sys.executable
            start_menu = os.path.join(os.environ.get('APPDATA', ''),
                                      'Microsoft', 'Windows', 'Start Menu', 'Programs', 'Inosys')

            os.makedirs(start_menu, exist_ok=True)
            shortcut_path = os.path.join(start_menu, "Primus Implant Report Generator.lnk")

            ps_script = f'''
    $WshShell = New-Object -comObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut("{shortcut_path}")
    $Shortcut.TargetPath = "{current_exe}"
    $Shortcut.WorkingDirectory = "{os.path.dirname(current_exe)}"
    $Shortcut.IconLocation = "{current_exe}"
    $Shortcut.Description = "Primus Implant Report Generator"
    $Shortcut.Save()
    '''

            result = subprocess.run(['powershell', '-Command', ps_script],
                                    capture_output=True, text=True)

            if result.returncode == 0:
                print("Start Menu shortcut created successfully")

        except Exception as e:
            print(f"Error creating Start Menu shortcut: {e}")

    def setup_user_installation(self) -> None:
        """Set up user-level installation on first run"""
        try:
            user_dir = get_user_app_directory()
            setup_marker = os.path.join(user_dir, '.setup_complete')

            # Check if setup has already been completed
            if os.path.exists(setup_marker):
                return

            print("Setting up user installation...")

            # Ensure directories exist
            self.ensure_user_directories()

            # Copy data files
            self.copy_data_files_to_user_directory()

            # Create shortcuts
            self.create_desktop_shortcut()
            self.create_start_menu_shortcut()

            # Create setup marker
            with open(setup_marker, 'w') as f:
                f.write(f"Setup completed on {datetime.now().isoformat()}\n")
                f.write(f"Version: {APP_VERSION}\n")

            print("User installation setup complete!")

        except Exception as e:
            print(f"Error during user installation setup: {e}")

    # Update the load_implant_data method to use user directory
    def load_implant_data(self) -> None:
        """Load implant data from CSV file"""
        csv_filename = self.get_data_file_path("Primus Implant List - Primus Implant List.csv")

        try:
            if not os.path.exists(csv_filename):
                raise FileNotFoundError(f"CSV file not found at {csv_filename}")

            self.implant_data = pd.read_csv(csv_filename)
            print(f"Implant data loaded successfully from {csv_filename}!")
            print(f"Total records: {len(self.implant_data)}")

            # Validate required columns
            required_columns = [
                'Implant Line', 'Implant Part No', 'Implant Diameter', 'Implant Length',
                'Guide Sleeve', 'Drill Length', 'Offset', 'Starter Drill',
                'Initial Drill 1', 'Initial Drill 2', 'Drill 1', 'Drill 2', 'Drill 3', 'Drill 4'
            ]

            missing_columns = [col for col in required_columns if col not in self.implant_data.columns]
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

    def bind_enter_keys(self) -> None:
        """Bind Enter key to appropriate actions based on current tab"""

        def on_enter_pressed(event):
            current_tab = self.notebook.get()
            if current_tab == "Add Implant":
                self.add_implants_to_plan()
            elif current_tab == "Generate Report":
                self.generate_pdf_report()

        # Bind to the main window
        self.bind('<Return>', on_enter_pressed)
        self.bind('<KP_Enter>', on_enter_pressed)  # Numpad Enter

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
            text="Primus Implant Report Generator",
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

        self.bind_enter_keys()

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
        """Check for updates in user directory"""
        checking_dialog = self.show_update_dialog("Checking for updates...", show_progress=True)

        update_thread = threading.Thread(target=self._user_update_check_worker, args=(checking_dialog,))
        update_thread.daemon = True
        update_thread.start()

    def _user_update_check_worker(self, checking_dialog) -> None:
        """User-level update check"""
        try:
            update_path = os.path.join(UPDATE_SERVER_PATH, UPDATE_FILENAME)

            if not os.path.exists(update_path):
                self.after(100, lambda: self._show_update_result(checking_dialog, "no_update"))
                return

            if getattr(sys, 'frozen', False):
                current_exe = sys.executable

                try:
                    update_stat = os.stat(update_path)
                    current_stat = os.stat(current_exe)

                    if update_stat.st_mtime <= current_stat.st_mtime:
                        self.after(100, lambda: self._show_update_result(checking_dialog, "up_to_date"))
                        return
                except:
                    pass

            self.after(100, lambda: self._show_update_result(checking_dialog, "update_available", update_path))

        except Exception as e:
            error_msg = f"Update check failed: {str(e)}"
            self.after(100, lambda: self._show_update_result(checking_dialog, "error", error_msg))

    def _update_check_worker(self, checking_dialog: ctk.CTkToplevel) -> None:
        """Worker thread for checking and downloading updates"""
        try:
            # Look for version-specific files first, then fall back to generic name
            version_specific_filename = f"Primus Implant Report Generator v{APP_VERSION}.exe"
            generic_filename = UPDATE_FILENAME

            version_specific_path = os.path.join(UPDATE_SERVER_PATH, version_specific_filename)
            generic_path = os.path.join(UPDATE_SERVER_PATH, generic_filename)

            update_path = None

            # Check for newer version files (look for versions higher than current)
            try:
                if os.path.exists(UPDATE_SERVER_PATH):
                    for filename in os.listdir(UPDATE_SERVER_PATH):
                        if filename.startswith("Primus Implant Report Generator v") and filename.endswith(
                                ".exe"):
                            # Extract version from filename
                            try:
                                version_part = filename.replace("Primus Implant Report Generator v", "").replace(
                                    ".exe", "")
                                # Simple version comparison (works for x.y.z format)
                                if self._is_newer_version(version_part, APP_VERSION):
                                    update_path = os.path.join(UPDATE_SERVER_PATH, filename)
                                    break
                            except:
                                continue
            except:
                pass

            # If no version-specific newer file found, check generic file
            if not update_path:
                if os.path.exists(generic_path):
                    update_path = generic_path
                else:
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

                if current_stat and update_stat.st_mtime <= current_stat.st_mtime and update_path == generic_path:
                    self.after(100, lambda: self._show_update_result(checking_dialog, "up_to_date"))
                    return
            except:
                pass  # If we can't compare, proceed with update

            # Update available, start download
            self.after(100, lambda: self._show_update_result(checking_dialog, "update_available", update_path))

        except Exception as e:
            error_msg = f"Update check failed: {str(e)}"
            self.after(100, lambda: self._show_update_result(checking_dialog, "error", error_msg))

    def _is_newer_version(self, version1: str, version2: str) -> bool:
        """Compare two version strings (e.g., '1.0.2' vs '1.0.1')"""
        try:
            v1_parts = [int(x) for x in version1.split('.')]
            v2_parts = [int(x) for x in version2.split('.')]

            # Pad shorter version with zeros
            max_len = max(len(v1_parts), len(v2_parts))
            v1_parts.extend([0] * (max_len - len(v1_parts)))
            v2_parts.extend([0] * (max_len - len(v2_parts)))

            return v1_parts > v2_parts
        except:
            return False

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
        """User-level update installation"""
        progress_dialog = self.show_update_dialog("Installing update...", show_progress=True)

        download_thread = threading.Thread(target=self._user_level_update_worker, args=(update_path, progress_dialog))
        download_thread.daemon = True
        download_thread.start()

    def _user_level_update_worker(self, update_path: str, progress_dialog) -> None:
        """Perform update in user directory without admin privileges"""
        try:
            if not getattr(sys, 'frozen', False):
                self.after(100, lambda: self._show_download_result(progress_dialog, "error",
                                                                   "Updates are only available for compiled executables."))
                return

            current_exe = sys.executable
            current_dir = os.path.dirname(current_exe)
            user_dir = get_user_app_directory()

            # Ensure we're running from user directory
            if not current_exe.startswith(user_dir):
                self.after(100, lambda: self._show_download_result(progress_dialog, "error",
                                                                   "Please run the application from the user installation to update."))
                return

            # Create backup
            backup_path = current_exe + ".backup"
            if os.path.exists(backup_path):
                os.remove(backup_path)
            shutil.copy2(current_exe, backup_path)

            # Copy new version to temporary location
            temp_update = os.path.join(current_dir, "update_temp.exe")
            shutil.copy2(update_path, temp_update)

            # Create update script
            update_script = self._create_user_update_script(current_exe, temp_update, backup_path)

            self.after(100, lambda: self._show_download_result(progress_dialog, "success", update_script))

        except Exception as e:
            error_msg = f"Update failed: {str(e)}"
            self.after(100, lambda: self._show_download_result(progress_dialog, "error", error_msg))

    def _create_user_update_script(self, current_exe: str, temp_update: str, backup_path: str) -> str:
        """Create user-level update script"""
        current_dir = os.path.dirname(current_exe)
        script_path = os.path.join(current_dir, "perform_update.bat")

        script_content = f'''@echo off
    echo Updating Primus Implant Report Generator...

    REM Wait for application to close
    timeout /t 3 /nobreak >nul

    REM Kill any remaining processes
    taskkill /f /im "{os.path.basename(current_exe)}" 2>nul

    REM Wait a moment more
    timeout /t 2 /nobreak >nul

    REM Perform the update
    echo Applying update...

    REM Replace the executable
    move "{temp_update}" "{current_exe}"

    if %errorlevel% equ 0 (
        echo Update successful!

        REM Clean up backup
        if exist "{backup_path}" del "{backup_path}"

        REM Update shortcuts to ensure they're current
        call :UpdateShortcuts

        echo Starting updated application...
        start "" "{current_exe}"

    ) else (
        echo Update failed, restoring backup...
        if exist "{backup_path}" (
            move "{backup_path}" "{current_exe}"
            echo Backup restored.
        )

        REM Clean up failed update file
        if exist "{temp_update}" del "{temp_update}"

        echo Starting original application...
        start "" "{current_exe}"
    )

    REM Clean up this script
    timeout /t 2 /nobreak >nul
    del "%~f0"
    exit /b

    :UpdateShortcuts
    REM Update desktop shortcut
    set "DESKTOP_SHORTCUT=%USERPROFILE%\\Desktop\\Primus Implant Report Generator.lnk"
    if exist "%DESKTOP_SHORTCUT%" (
        powershell -Command "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%DESKTOP_SHORTCUT%'); $Shortcut.TargetPath = '{current_exe}'; $Shortcut.WorkingDirectory = '{current_dir}'; $Shortcut.IconLocation = '{current_exe}'; $Shortcut.Save()"
    )

    REM Update Start Menu shortcut
    set "START_MENU_SHORTCUT=%APPDATA%\\Microsoft\\Windows\\Start Menu\\Programs\\Inosys\\Primus Implant Report Generator.lnk"
    if exist "%START_MENU_SHORTCUT%" (
        powershell -Command "$WshShell = New-Object -comObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%START_MENU_SHORTCUT%'); $Shortcut.TargetPath = '{current_exe}'; $Shortcut.WorkingDirectory = '{current_dir}'; $Shortcut.IconLocation = '{current_exe}'; $Shortcut.Save()"
    )
    exit /b
    '''

        with open(script_path, 'w') as f:
            f.write(script_content)

        return script_path

    def get_update_server_info(self) -> dict:
        """Get update server information - can be customized for different deployment methods"""
        return {
            'type': 'network_share',  # or 'web', 'local'
            'path': UPDATE_SERVER_PATH,
            'filename': UPDATE_FILENAME,
            'version_file': 'version.txt'
        }

    def get_update_info(self) -> dict:
        """Get comprehensive update information"""
        try:
            update_manifest_path = os.path.join(UPDATE_SERVER_PATH, 'update_manifest.json')

            if os.path.exists(update_manifest_path):
                with open(update_manifest_path, 'r') as f:
                    return json.load(f)
            else:
                # Fallback to basic file checking
                update_exe_path = os.path.join(UPDATE_SERVER_PATH, UPDATE_FILENAME)
                if os.path.exists(update_exe_path):
                    stat_info = os.stat(update_exe_path)
                    return {
                        "version": "Unknown",
                        "build_date": datetime.fromtimestamp(stat_info.st_mtime).isoformat(),
                        "executable": UPDATE_FILENAME,
                        "size": stat_info.st_size
                    }

            return None
        except Exception as e:
            print(f"Error getting update info: {e}")
            return None

    def check_update_server_access(self) -> bool:
        """Enhanced update server access check"""
        try:
            # Check if primary network path exists
            network_path = r"\\CDIMANQ30\Creoman-Active\CADCAM\Software\Primus Implant Report Generator\UserUpdates"
            if os.path.exists(network_path):
                self.log_update_activity(f"Network update server accessible: {network_path}")
                return True

            # Check fallback local path
            local_path = r"C:\Shared\PrimusUpdates"
            if os.path.exists(local_path):
                self.log_update_activity(f"Local update server accessible: {local_path}")
                return True

            # No update server accessible
            self.log_update_activity("No update servers accessible")
            return False

        except Exception as e:
            self.log_update_activity(f"Error checking update server access: {e}")
            return False

    def check_update_server_access(self) -> bool:
        """Check if we can access the update server"""
        try:
            server_info = self.get_update_server_info()

            if server_info['type'] == 'network_share':
                return os.path.exists(server_info['path'])
            elif server_info['type'] == 'web':
                # Add web-based update check here if needed
                pass
            elif server_info['type'] == 'local':
                return os.path.exists(server_info['path'])

            return False
        except:
            return False

    def _show_download_result(self, progress_dialog, result: str, data: str) -> None:
        """Show update result with user-friendly messages"""
        progress_dialog.destroy()

        if result == "error":
            messagebox.showerror("Update Error",
                                 f"{data}\n\nYou can manually download the latest version from the network location.")
        elif result == "success":
            response = messagebox.askyesno(
                "Update Ready",
                "Update is ready to install. The application will close and restart automatically.\n\n"
                "Click Yes to proceed with the update, or No to update later."
            )
            if response:
                # Close application and run update script
                subprocess.Popen(data, shell=True)
                self.quit()

    # Optional: Add a method to check for updates on startup
    def check_for_updates_on_startup(self) -> None:
        """Optionally check for updates when the application starts"""
        # This could be enabled via a setting
        if hasattr(self, 'auto_check_updates') and self.auto_check_updates:
            # Check in background without showing dialog
            threading.Thread(target=self._silent_update_check, daemon=True).start()

    def _silent_update_check(self) -> None:
        """Silent update check for startup"""
        try:
            if not self.check_update_server_access():
                return

            update_path = os.path.join(UPDATE_SERVER_PATH, UPDATE_FILENAME)
            if not os.path.exists(update_path):
                return

            if getattr(sys, 'frozen', False):
                current_exe = sys.executable

                try:
                    update_stat = os.stat(update_path)
                    current_stat = os.stat(current_exe)

                    if update_stat.st_mtime > current_stat.st_mtime:
                        # Show update notification in main thread
                        self.after(100, self._show_update_notification)

                except Exception:
                    pass

        except Exception as e:
            print(f"Silent update check failed: {e}")

    def _show_update_notification(self) -> None:
        """Show update notification to user"""
        response = messagebox.askyesno(
            "Update Available",
            "A new version of Primus Implant Report Generator is available.\n\n"
            "Would you like to update now?"
        )

        if response:
            self.check_for_updates()

    # Add logging for better troubleshooting
    def log_update_activity(self, message: str) -> None:
        """Log update activities for troubleshooting"""
        try:
            user_dir = get_user_app_directory()
            log_file = os.path.join(user_dir, 'logs', 'updates.log')

            os.makedirs(os.path.dirname(log_file), exist_ok=True)

            with open(log_file, 'a', encoding='utf-8') as f:
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                f.write(f"[{timestamp}] {message}\n")

        except Exception as e:
            print(f"Failed to write log: {e}")

    def log_window_activity(self, message: str, level: str = "INFO") -> None:
        """Log window memory activities for debugging"""
        try:
            user_dir = get_user_app_directory()
            log_dir = os.path.join(user_dir, 'logs')
            log_file = os.path.join(log_dir, 'window_memory.log')

            os.makedirs(log_dir, exist_ok=True)

            with open(log_file, 'a', encoding='utf-8') as f:
                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                f.write(f"[{timestamp}] [{level}] {message}\n")

            # Also print to console for immediate feedback
            print(f"[WINDOW_MEMORY] [{level}] {message}")

        except Exception as e:
            print(f"Failed to write window memory log: {e}")

    # Enhanced update check with better error handling
    def _user_update_check_worker(self, checking_dialog) -> None:
        """Enhanced user-level update check with logging"""
        try:
            self.log_update_activity("Starting update check...")

            # Check server access first
            if not self.check_update_server_access():
                self.log_update_activity("Update server not accessible")
                self.after(100, lambda: self._show_update_result(checking_dialog, "no_server"))
                return

            update_path = os.path.join(UPDATE_SERVER_PATH, UPDATE_FILENAME)
            self.log_update_activity(f"Checking for update at: {update_path}")

            if not os.path.exists(update_path):
                self.log_update_activity("No update file found")
                self.after(100, lambda: self._show_update_result(checking_dialog, "no_update"))
                return

            if getattr(sys, 'frozen', False):
                current_exe = sys.executable
                self.log_update_activity(f"Current executable: {current_exe}")

                try:
                    update_stat = os.stat(update_path)
                    current_stat = os.stat(current_exe)

                    update_time = datetime.fromtimestamp(update_stat.st_mtime)
                    current_time = datetime.fromtimestamp(current_stat.st_mtime)

                    self.log_update_activity(f"Update file time: {update_time}")
                    self.log_update_activity(f"Current file time: {current_time}")

                    if update_stat.st_mtime <= current_stat.st_mtime:
                        self.log_update_activity("Current version is up to date")
                        self.after(100, lambda: self._show_update_result(checking_dialog, "up_to_date"))
                        return

                except Exception as e:
                    self.log_update_activity(f"Error comparing file times: {e}")

            self.log_update_activity("Update available, prompting user")
            self.after(100, lambda: self._show_update_result(checking_dialog, "update_available", update_path))

        except Exception as e:
            error_msg = f"Update check failed: {str(e)}"
            self.log_update_activity(error_msg)
            self.after(100, lambda: self._show_update_result(checking_dialog, "error", error_msg))

    def _show_update_result(self, checking_dialog, result: str, data: str = None) -> None:
        """Enhanced update result handling"""
        checking_dialog.destroy()

        if result == "no_update":
            messagebox.showinfo("No Updates", "No updates are available at this time.")
        elif result == "up_to_date":
            messagebox.showinfo("Up to Date", "You are running the latest version.")
        elif result == "no_server":
            messagebox.showwarning("Update Server",
                                   "Cannot connect to update server. Please check your network connection or contact IT support.")
        elif result == "error":
            messagebox.showerror("Update Error",
                                 f"{data}\n\nPlease try again later or contact support if the problem persists.")
        elif result == "update_available":
            response = messagebox.askyesno("Update Available",
                                           "A new version is available. Would you like to download and install it now?\n\n"
                                           "The application will restart after the update.\n\n"
                                           "Note: This update will install without requiring administrator privileges.")
            if response:
                self._download_and_install_update(data)

    # Method to create a settings file for update preferences
    def create_update_settings(self) -> None:
        """Create update settings file"""
        try:
            user_dir = get_user_app_directory()
            settings_file = os.path.join(user_dir, 'update_settings.json')

            default_settings = {
                "auto_check_on_startup": False,
                "check_interval_hours": 24,
                "update_server_type": "network_share",
                "custom_server_path": "",
                "last_check": None,
                "skip_version": None
            }

            if not os.path.exists(settings_file):
                import json
                with open(settings_file, 'w') as f:
                    json.dump(default_settings, f, indent=2)

        except Exception as e:
            print(f"Error creating update settings: {e}")

    # Method to read update settings
    def load_update_settings(self) -> dict:
        """Load update settings"""
        try:
            user_dir = get_user_app_directory()
            settings_file = os.path.join(user_dir, 'update_settings.json')

            if os.path.exists(settings_file):
                import json
                with open(settings_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading update settings: {e}")

        # Return defaults if file doesn't exist or can't be read
        return {
            "auto_check_on_startup": False,
            "check_interval_hours": 24,
            "update_server_type": "network_share",
            "custom_server_path": "",
            "last_check": None,
            "skip_version": None
        }

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
echo Updating Primus Implant Report Generator...
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
        """Enhanced Generate Report tab with case notes and print preview"""
        tab: ctk.CTkFrame = self.notebook.tab("Generate Report")

        # Main scrollable frame
        main_scrollable = ctk.CTkScrollableFrame(tab, fg_color=INOSYS_COLORS["background_secondary"])
        main_scrollable.pack(fill="both", expand=True, padx=10, pady=10)

        # Title
        title_label: ctk.CTkLabel = ctk.CTkLabel(
            main_scrollable,
            text="Generate PDF Report",
            font=ctk.CTkFont(size=18, weight="bold"),
            text_color=INOSYS_COLORS["light_blue"]
        )
        title_label.pack(pady=20)

        # Doctor information frame
        info_frame: ctk.CTkFrame = ctk.CTkFrame(main_scrollable, fg_color=INOSYS_COLORS["background_tertiary"])
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

        # Case Notes section
        notes_frame: ctk.CTkFrame = ctk.CTkFrame(main_scrollable, fg_color=INOSYS_COLORS["background_tertiary"])
        notes_frame.pack(fill="x", padx=20, pady=10)

        ctk.CTkLabel(
            notes_frame,
            text="Case Notes (Optional):",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=INOSYS_COLORS["text_primary"]
        ).pack(anchor="w", padx=10, pady=(10, 5))

        # Create a frame for the text widget to control colors better
        text_container = ctk.CTkFrame(notes_frame, fg_color=INOSYS_COLORS["background_secondary"])
        text_container.pack(fill="x", padx=10, pady=(0, 10))

        # Use tkinter Text widget for better functionality
        self.case_notes_text = tk.Text(
            text_container,
            height=4,
            wrap=tk.WORD,
            bg=INOSYS_COLORS["background_secondary"],
            fg=INOSYS_COLORS["text_primary"],
            insertbackground=INOSYS_COLORS["text_primary"],
            selectbackground=INOSYS_COLORS["medium_blue"],
            selectforeground=INOSYS_COLORS["white"],
            relief="flat",
            bd=5,
            font=("Segoe UI", 10)
        )
        self.case_notes_text.pack(fill="x", padx=5, pady=5)

        # Placeholder text functionality
        self.setup_notes_placeholder()

        # Button frame
        button_frame: ctk.CTkFrame = ctk.CTkFrame(main_scrollable, fg_color="transparent")
        button_frame.pack(pady=20)

        # Print Preview button
        preview_button: ctk.CTkButton = ctk.CTkButton(
            button_frame,
            text="View Report",
            command=self.show_print_preview,
            height=50,
            width=180,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=INOSYS_COLORS["medium_blue"],
            hover_color=INOSYS_COLORS["light_blue"],
            text_color=INOSYS_COLORS["white"]
        )
        preview_button.pack(side="left", padx=10)

        # Generate PDF button
        generate_button: ctk.CTkButton = ctk.CTkButton(
            button_frame,
            text="Save Report",
            command=self.generate_pdf_report,
            height=50,
            width=180,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=INOSYS_COLORS["light_blue"],
            hover_color=INOSYS_COLORS["medium_blue"],
            text_color=INOSYS_COLORS["white"]
        )
        generate_button.pack(side="left", padx=10)

    def setup_notes_placeholder(self) -> None:
        """Setup placeholder text for case notes"""
        placeholder_text = "Enter any special instructions, patient considerations, or case-specific notes here..."

        def on_focus_in(event):
            if self.case_notes_text.get("1.0", tk.END).strip() == placeholder_text:
                self.case_notes_text.delete("1.0", tk.END)
                self.case_notes_text.config(fg=INOSYS_COLORS["text_primary"])

        def on_focus_out(event):
            if not self.case_notes_text.get("1.0", tk.END).strip():
                self.case_notes_text.insert("1.0", placeholder_text)
                self.case_notes_text.config(fg=INOSYS_COLORS["text_secondary"])

        # Set initial placeholder
        self.case_notes_text.insert("1.0", placeholder_text)
        self.case_notes_text.config(fg=INOSYS_COLORS["text_secondary"])

        # Bind events
        self.case_notes_text.bind("<FocusIn>", on_focus_in)
        self.case_notes_text.bind("<FocusOut>", on_focus_out)

    def get_case_notes(self) -> str:
        """Get case notes, excluding placeholder text"""
        notes = self.case_notes_text.get("1.0", tk.END).strip()
        placeholder_text = "Enter any special instructions, patient considerations, or case-specific notes here..."

        if notes == placeholder_text or not notes:
            return ""
        return notes

    def get_window_settings_file(self) -> str:
        """Get the path to the window settings file"""
        try:
            user_dir = get_user_app_directory()
            self.log_window_activity(f"User directory: {user_dir}")

            # Debug: Show environment variables
            self.log_window_activity(f"USERNAME: {os.environ.get('USERNAME', 'NOT_SET')}")
            self.log_window_activity(f"USERPROFILE: {os.environ.get('USERPROFILE', 'NOT_SET')}")
            self.log_window_activity(f"LOCALAPPDATA: {os.environ.get('LOCALAPPDATA', 'NOT_SET')}")

            # Ensure directory exists
            os.makedirs(user_dir, exist_ok=True)
            self.log_window_activity(f"Directory created/verified: {user_dir}")

            settings_file = os.path.join(user_dir, 'window_settings.json')
            self.log_window_activity(f"Settings file path: {settings_file}")

            return settings_file

        except Exception as e:
            self.log_window_activity(f"Error getting settings file path: {e}", "ERROR")
            return None

    def show_print_preview(self) -> None:
        """Generate and show print preview"""
        if not self.implant_plans:
            messagebox.showerror("Error", "No implant plans to preview!")
            return

        try:
            # Get form data
            doctor_name: str = self.doctor_name_entry.get() or "Dr. [Name]"
            patient_name: str = self.patient_name_entry.get() or "[Patient Name]"
            case_number: str = self.case_number_entry.get() or "[Case Number]"
            case_notes: str = self.get_case_notes()

            # Create temporary PDF
            temp_dir = tempfile.gettempdir()
            temp_filename = os.path.join(temp_dir, f"primus_preview_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")

            # Generate preview PDF
            self.create_pdf_report(temp_filename, doctor_name, patient_name, case_number, case_notes, is_preview=True)

            # Open with default PDF viewer
            if sys.platform.startswith('win'):
                os.startfile(temp_filename)
            elif sys.platform.startswith('darwin'):  # macOS
                subprocess.run(['open', temp_filename])
            else:  # Linux
                subprocess.run(['xdg-open', temp_filename])

            # Show preview dialog
            self.show_preview_dialog(temp_filename)

        except Exception as e:
            messagebox.showerror("Preview Error", f"Failed to generate preview: {str(e)}")

    def show_preview_dialog(self, temp_filename: str) -> None:
        """Show preview dialog with options"""
        preview_dialog = ctk.CTkToplevel(self)
        preview_dialog.title("Print Preview")
        preview_dialog.geometry("400x200")
        preview_dialog.resizable(False, False)
        preview_dialog.configure(fg_color=INOSYS_COLORS["background_primary"])

        # Center the dialog
        preview_dialog.transient(self)
        preview_dialog.grab_set()

        # Main frame
        main_frame = ctk.CTkFrame(preview_dialog, fg_color=INOSYS_COLORS["background_secondary"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Message
        message_label = ctk.CTkLabel(
            main_frame,
            text="Preview opened in your default PDF viewer.\n\nHow would you like to proceed?",
            font=ctk.CTkFont(size=12),
            text_color=INOSYS_COLORS["text_primary"]
        )
        message_label.pack(pady=20)

        # Button frame
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.pack(pady=10)

        # Save As button
        save_button = ctk.CTkButton(
            button_frame,
            text="Save As...",
            command=lambda: self.save_preview_as(temp_filename, preview_dialog),
            width=100,
            fg_color=INOSYS_COLORS["light_blue"],
            hover_color=INOSYS_COLORS["medium_blue"]
        )
        save_button.pack(side="left", padx=10)

        # Print button (opens print dialog)
        print_button = ctk.CTkButton(
            button_frame,
            text="Print",
            command=lambda: self.print_preview(temp_filename, preview_dialog),
            width=100,
            fg_color=INOSYS_COLORS["medium_blue"],
            hover_color=INOSYS_COLORS["light_blue"]
        )
        print_button.pack(side="left", padx=10)

        # Close button
        close_button = ctk.CTkButton(
            button_frame,
            text="Close",
            command=lambda: self.close_preview(temp_filename, preview_dialog),
            width=100,
            fg_color=INOSYS_COLORS["dark_blue"],
            hover_color=INOSYS_COLORS["medium_blue"]
        )
        close_button.pack(side="left", padx=10)

    def save_preview_as(self, temp_filename: str, dialog: ctk.CTkToplevel) -> None:
        """Save preview as permanent file"""
        try:
            case_number = self.case_number_entry.get() or "Case"
            filename = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                title="Save Report As",
                initialfile=f"Primus_Report_{case_number}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            )

            if filename:
                shutil.copy2(temp_filename, filename)
                messagebox.showinfo("Success", f"Report saved successfully!\nSaved as: {filename}")
                dialog.destroy()

        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save report: {str(e)}")

    def print_preview(self, temp_filename: str, dialog: ctk.CTkToplevel) -> None:
        """Send preview to printer"""
        try:
            if sys.platform.startswith('win'):
                # Use Windows print command
                subprocess.run(
                    f'rundll32 advapi32.dll,ProcessIdleTasks & ping 127.0.0.1 -n 2 > nul & "{temp_filename}"',
                    shell=True)
            else:
                messagebox.showinfo("Print", "Please use your PDF viewer's print function to print the document.")

            dialog.destroy()

        except Exception as e:
            messagebox.showerror("Print Error", f"Failed to print: {str(e)}")

    def close_preview(self, temp_filename: str, dialog: ctk.CTkToplevel) -> None:
        """Close preview and cleanup"""
        try:
            # Clean up temporary file
            if os.path.exists(temp_filename):
                os.remove(temp_filename)
        except:
            pass  # Ignore cleanup errors

        dialog.destroy()

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
            details_text: str = (
                f"Tooth {plan['tooth_number']}: {plan['implant_line']} - "
                f"{plan['diameter']}mm × {plan['length']}mm (Offset: {plan['offset']}mm)"
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
        """Enhanced PDF report generation with case notes"""
        if not self.implant_plans:
            messagebox.showerror("Error", "No implant plans to generate report!")
            return

        # Get report information
        doctor_name: str = self.doctor_name_entry.get() or "Dr. [Name]"
        patient_name: str = self.patient_name_entry.get() or "[Patient Name]"
        case_number: str = self.case_number_entry.get() or "[Case Number]"
        case_notes: str = self.get_case_notes()

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
            self.create_pdf_report(filename, doctor_name, patient_name, case_number, case_notes)
            messagebox.showinfo("Success", f"Report generated successfully!\nSaved as: {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate report: {str(e)}")

    def create_pdf_report(self, filename: str, doctor_name: str, patient_name: str, case_number: str,
                          case_notes: str = "", is_preview: bool = False) -> None:
        """Enhanced PDF report creation with case notes"""
        doc: SimpleDocTemplate = SimpleDocTemplate(
            filename,
            pagesize=letter,
            topMargin=0.5 * inch,
            bottomMargin=0.5 * inch,
            leftMargin=0.5 * inch,
            rightMargin=0.5 * inch,
            title="Primus Implant Report" + (" - Preview" if is_preview else "")
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
            story.append(Paragraph("PRIMUS IMPLANT SURGICAL DRILLING PROTOCOL", title_style))

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

        # Main implant data table with better column sizing
        implant_data = [
            ["Tooth", "Part Number", "Dia.", "Len.", "Offset", "Guide Sleeve", "Drill Length", "Drilling Sequence"]]

        for i, plan in enumerate(sorted_plans):
            # Create drilling sequence string with surgical approach instructions
            # Handle 'x' values in drilling sequence
            def format_drill_value(value):
                return "N/A" if str(value).lower().strip() == 'x' else f"{value}mm"

            # Determine approach text
            is_flapless = plan.get('surgical_approach', 'flapless') == 'flapless'
            approach_instruction = "Tissue punch → Drill to bone → clear tissue" if is_flapless else "Open flap and reflect tissue prior to seating surgical guide"

            # Create drilling sequence with approach instruction and consistent font
            drill_sequence = f"<b>{approach_instruction}</b><br/>" \
                             f"Start: {format_drill_value(plan['implant_data']['Starter Drill'])} → Init1: {format_drill_value(plan['implant_data']['Initial Drill 1'])} → Init2: {format_drill_value(plan['implant_data']['Initial Drill 2'])}<br/>" \
                             f"D1: {format_drill_value(plan['implant_data']['Drill 1'])} → D2: {format_drill_value(plan['implant_data']['Drill 2'])} → D3: {format_drill_value(plan['implant_data']['Drill 3'])} → D4: {format_drill_value(plan['implant_data']['Drill 4'])}"

            implant_data.append([
                str(plan['tooth_number']),
                plan['implant_data']['Implant Part No'],
                f"{plan['diameter']}mm",
                f"{plan['length']}mm",
                f"{plan['offset']}mm",
                plan['implant_data']['Guide Sleeve'],
                f"{plan['implant_data']['Drill Length']}mm",
                Paragraph(drill_sequence, ParagraphStyle('DrillSeq',
                                                         parent=styles['Normal'],
                                                         fontSize=7,
                                                         fontName='Helvetica',
                                                         leading=8))
            ])

        # Better balanced column widths - adjusted for overflow issues
        col_widths = [
            0.35 * inch,  # Tooth - slightly smaller
            0.7 * inch,  # Part Number - fits PBF4010S
            0.35 * inch,  # Diameter
            0.45 * inch,  # Length - widened
            0.4 * inch,  # Offset
            0.8 * inch,  # Guide Sleeve - increased for CGSC-5304
            0.65 * inch,  # Drill Length - widened
            2.9 * inch  # Drilling Sequence - slightly reduced to compensate
        ]

        implant_table: Table = Table(implant_data, colWidths=col_widths, repeatRows=1)

        implant_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('FONTSIZE', (0, 1), (6, -1), 7),  # Smaller font for data to prevent overflow
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('BACKGROUND', (0, 0), (-1, 0), colors.Color(30 / 255, 58 / 255, 138 / 255)),  # Dark blue header
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            # Better text wrapping and overflow handling
            ('WORDWRAP', (1, 1), (1, -1), True),  # Part Number
            ('WORDWRAP', (5, 1), (5, -1), True),  # Guide Sleeve
            ('WORDWRAP', (7, 1), (7, -1), True),  # Drilling Sequence
            # Prevent text overflow
            ('OVERFLOW', (0, 0), (-1, -1), 'CLIP'),
        ]))

        story.append(implant_table)
        story.append(Spacer(1, 15))

        # Add case notes if they exist
        if case_notes:
            story.append(Spacer(1, 15))

            # Case Notes section
            story.append(Paragraph("CASE NOTES", header_style))

            # Create case notes box
            case_notes_style = ParagraphStyle(
                'CaseNotes',
                parent=styles['Normal'],
                fontSize=10,
                leading=12,
                leftIndent=10,
                rightIndent=10,
                spaceBefore=8,
                spaceAfter=8,
                borderWidth=1,
                borderColor=colors.Color(14 / 255, 165 / 255, 233 / 255),
                borderPadding=10,
                backColor=colors.Color(248 / 255, 249 / 255, 250 / 255)
            )

            story.append(Paragraph(case_notes, case_notes_style))
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
                "• Speeds: 300 RPM tissue punch, 1200 RPM cortical perforator,\n"
                "  800 RPM shaping drills\n"
                "• Apply light pressure - let drill do the work\n"
                "• Use copious irrigation (minimum 50ml/min)"
            ],
            ["POST-DRILLING VERIFICATION", "IMPORTANT NOTES"],
            [
                "• Irrigate osteotomy thoroughly\n"
                "• Check final depth and angulation\n"
                "• Verify diameter with sizing gauge\n"
                "• Ensure proper guide sleeve function\n"
                "• Proceed with implant placement protocol\n"
                "• Implant placement speed & torque: 20 RPM 35 Ncm ",

                "• Follow manufacturer drilling guidelines\n"
                "• Maintain sterile technique throughout\n"
                "• Account for offset measurements\n"
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
        footer_text: str = f"""<b>Report Generated:</b> {datetime.now().strftime("%B %d, %Y at %I:%M %p")} | <b>Software:</b> Primus Implant Report Generator v{APP_VERSION}"""

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

    def save_window_geometry(self) -> None:
        """Save window size and position with extensive logging"""
        try:
            self.log_window_activity("=== SAVING WINDOW GEOMETRY ===")

            # Get settings file path
            settings_file = self.get_window_settings_file()
            if not settings_file:
                self.log_window_activity("Could not determine settings file path", "ERROR")
                return

            # Update widget info to get accurate measurements
            self.update_idletasks()
            self.log_window_activity("Updated idle tasks")

            # Get all window information
            try:
                geometry = self.winfo_geometry()
                self.log_window_activity(f"Raw geometry string: {geometry}")
            except Exception as e:
                self.log_window_activity(f"Error getting geometry: {e}", "ERROR")
                geometry = "900x950+100+100"

            try:
                width = self.winfo_width()
                height = self.winfo_height()
                x = self.winfo_x()
                y = self.winfo_y()
                self.log_window_activity(f"Individual values - Width: {width}, Height: {height}, X: {x}, Y: {y}")
            except Exception as e:
                self.log_window_activity(f"Error getting individual geometry values: {e}", "ERROR")
                width, height, x, y = 900, 950, 100, 100

            try:
                window_state = self.state()
                self.log_window_activity(f"Window state: {window_state}")
            except Exception as e:
                self.log_window_activity(f"Error getting window state: {e}", "ERROR")
                window_state = "normal"

            # Validate values
            if width < 300:
                self.log_window_activity(f"Width too small ({width}), using default", "WARN")
                width = 900
            if height < 200:
                self.log_window_activity(f"Height too small ({height}), using default", "WARN")
                height = 950

            # Create settings dictionary
            settings = {
                'geometry': geometry,
                'width': width,
                'height': height,
                'x': x,
                'y': y,
                'maximized': window_state == 'zoomed',
                'state': window_state,
                'saved_at': datetime.now().isoformat(),
                'version': '1.0.4'
            }

            self.log_window_activity(f"Settings to save: {settings}")

            # Write to file
            try:
                # First, try to read existing file to see if we can write
                if os.path.exists(settings_file):
                    with open(settings_file, 'r') as f:
                        old_settings = json.load(f)
                    self.log_window_activity(f"Previous settings: {old_settings}")

                # Write new settings
                with open(settings_file, 'w') as f:
                    json.dump(settings, f, indent=2)

                self.log_window_activity(f"Settings saved successfully to: {settings_file}")

                # Verify the file was written correctly
                with open(settings_file, 'r') as f:
                    verification = json.load(f)
                self.log_window_activity(f"Verification read: {verification}")

            except Exception as e:
                self.log_window_activity(f"Error writing settings file: {e}", "ERROR")

            self.log_window_activity("=== SAVE COMPLETE ===")

        except Exception as e:
            self.log_window_activity(f"Unexpected error in save_window_geometry: {e}", "ERROR")
            import traceback
            self.log_window_activity(f"Traceback: {traceback.format_exc()}", "ERROR")

    def load_window_geometry(self) -> None:
        """Load and apply saved window size and position with extensive logging"""
        try:
            self.log_window_activity("=== LOADING WINDOW GEOMETRY ===")

            # Get settings file path
            settings_file = self.get_window_settings_file()
            if not settings_file:
                self.log_window_activity("Could not determine settings file path, using defaults", "WARN")
                self.geometry("900x950")
                return

            # Check if settings file exists
            if not os.path.exists(settings_file):
                self.log_window_activity(f"Settings file does not exist: {settings_file}")
                self.log_window_activity("Using default geometry: 900x950")
                self.geometry("900x950")
                return

            self.log_window_activity(f"Loading settings from: {settings_file}")

            # Read settings file
            try:
                with open(settings_file, 'r') as f:
                    settings = json.load(f)
                self.log_window_activity(f"Settings loaded: {settings}")
            except Exception as e:
                self.log_window_activity(f"Error reading settings file: {e}", "ERROR")
                self.geometry("900x950")
                return

            # Validate settings
            if not isinstance(settings, dict):
                self.log_window_activity("Settings is not a dictionary, using defaults", "ERROR")
                self.geometry("900x950")
                return

            # Try different methods to apply geometry
            geometry_applied = False

            # Method 1: Use geometry string
            if 'geometry' in settings and settings['geometry']:
                try:
                    geometry_str = settings['geometry']
                    self.log_window_activity(f"Attempting to apply geometry string: {geometry_str}")
                    self.geometry(geometry_str)

                    # Verify it was applied
                    self.update_idletasks()
                    current_geometry = self.winfo_geometry()
                    self.log_window_activity(f"Geometry applied, current: {current_geometry}")
                    geometry_applied = True

                except Exception as e:
                    self.log_window_activity(f"Failed to apply geometry string: {e}", "ERROR")

            # Method 2: Use individual values
            if not geometry_applied:
                try:
                    width = settings.get('width', 900)
                    height = settings.get('height', 950)
                    x = settings.get('x', 100)
                    y = settings.get('y', 100)

                    self.log_window_activity(f"Attempting individual geometry - W:{width} H:{height} X:{x} Y:{y}")

                    # Validate individual values
                    if width < 300 or width > 5000:
                        self.log_window_activity(f"Invalid width {width}, using 900", "WARN")
                        width = 900
                    if height < 200 or height > 3000:
                        self.log_window_activity(f"Invalid height {height}, using 950", "WARN")
                        height = 950
                    if x < -1000 or x > 5000:
                        self.log_window_activity(f"Invalid x position {x}, using 100", "WARN")
                        x = 100
                    if y < -100 or y > 3000:
                        self.log_window_activity(f"Invalid y position {y}, using 100", "WARN")
                        y = 100

                    geometry_str = f"{width}x{height}+{x}+{y}"
                    self.log_window_activity(f"Constructed geometry string: {geometry_str}")
                    self.geometry(geometry_str)

                    # Verify
                    self.update_idletasks()
                    current_geometry = self.winfo_geometry()
                    self.log_window_activity(f"Individual geometry applied, current: {current_geometry}")
                    geometry_applied = True

                except Exception as e:
                    self.log_window_activity(f"Failed to apply individual geometry: {e}", "ERROR")

            # Method 3: Fallback to defaults
            if not geometry_applied:
                self.log_window_activity("All geometry methods failed, using defaults", "WARN")
                self.geometry("900x950")

            # Handle maximized state
            try:
                if settings.get('maximized', False):
                    self.log_window_activity("Window should be maximized")
                    # Use after() to ensure geometry is applied first
                    self.after(200, lambda: self.apply_maximized_state())
                else:
                    self.log_window_activity("Window should be normal size")
            except Exception as e:
                self.log_window_activity(f"Error handling maximized state: {e}", "ERROR")

            self.log_window_activity("=== LOAD COMPLETE ===")

        except Exception as e:
            self.log_window_activity(f"Unexpected error in load_window_geometry: {e}", "ERROR")
            import traceback
            self.log_window_activity(f"Traceback: {traceback.format_exc()}", "ERROR")
            # Fallback to default
            self.geometry("900x950")

    def apply_maximized_state(self) -> None:
        """Apply maximized state with logging"""
        try:
            self.log_window_activity("Applying maximized state")
            current_state = self.state()
            self.log_window_activity(f"Current state before maximizing: {current_state}")

            self.state('zoomed')

            # Verify
            self.update_idletasks()
            new_state = self.state()
            self.log_window_activity(f"State after maximizing: {new_state}")

        except Exception as e:
            self.log_window_activity(f"Error applying maximized state: {e}", "ERROR")

    def apply_maximized_state(self) -> None:
        """Apply maximized state with logging"""
        try:
            self.log_window_activity("Applying maximized state")
            current_state = self.state()
            self.log_window_activity(f"Current state before maximizing: {current_state}")

            self.state('zoomed')

            # Verify
            self.update_idletasks()
            new_state = self.state()
            self.log_window_activity(f"State after maximizing: {new_state}")
        except Exception as e:
            self.log_window_activity(f"Error applying maximized state: {e}", "ERROR")
    def apply_individual_geometry(self, settings: dict) -> None:
        """Apply geometry using individual width/height/x/y values"""
        try:
            width = settings.get('width', 900)
            height = settings.get('height', 950)
            x = settings.get('x', 100)
            y = settings.get('y', 100)

            # Validate values are reasonable
            if width < 400 or width > 3840:  # Min 400px, max 4K width
                width = 900
            if height < 300 or height > 2160:  # Min 300px, max 4K height
                height = 950
            if x < -100 or x > 2560:  # Reasonable screen positions
                x = 100
            if y < -100 or y > 1440:
                y = 100

            geometry_string = f"{width}x{height}+{x}+{y}"
            self.geometry(geometry_string)
            print(f"Applied individual geometry: {geometry_string}")

        except Exception as e:
            print(f"Error applying individual geometry: {e}")
            self.geometry("900x950")

    def on_window_close(self) -> None:
        """Handle window close event with logging"""
        self.log_window_activity("=== WINDOW CLOSING ===")
        try:
            self.save_window_geometry()
            self.log_window_activity("Window geometry saved on close")
        except Exception as e:
            self.log_window_activity(f"Error saving geometry on close: {e}", "ERROR")

        try:
            self.quit()
        except Exception as e:
            self.log_window_activity(f"Error during quit: {e}", "ERROR")

    def test_window_memory(self) -> None:
        """Test function to verify window memory is working"""
        try:
            self.log_window_activity("=== TESTING WINDOW MEMORY ===")

            # Test save
            self.log_window_activity("Testing save...")
            self.save_window_geometry()

            # Test load
            self.log_window_activity("Testing load...")
            settings_file = self.get_window_settings_file()

            if os.path.exists(settings_file):
                with open(settings_file, 'r') as f:
                    settings = json.load(f)
                self.log_window_activity(f"Test read successful: {settings}")
            else:
                self.log_window_activity("Test failed: settings file not found", "ERROR")

            self.log_window_activity("=== TEST COMPLETE ===")

        except Exception as e:
            self.log_window_activity(f"Test failed: {e}", "ERROR")

    def open_window_log(self) -> None:
        """Open the window memory log file"""
        try:
            user_dir = get_user_app_directory()
            log_file = os.path.join(user_dir, 'logs', 'window_memory.log')

            if os.path.exists(log_file):
                if sys.platform.startswith('win'):
                    os.startfile(log_file)
                else:
                    subprocess.run(['open', log_file])
            else:
                messagebox.showinfo("Log File", f"Log file not found at:\n{log_file}")

        except Exception as e:
            messagebox.showerror("Error", f"Could not open log file: {e}")

    # Update your create_menu method to include the test option
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
        help_menu.add_separator()
        help_menu.add_command(label="Test Window Memory", command=self.show_window_memory_test)

    def show_window_memory_test(self) -> None:
        """Show window memory test dialog"""
        test_dialog = ctk.CTkToplevel(self)
        test_dialog.title("Window Memory Test")
        test_dialog.geometry("400x300")
        test_dialog.configure(fg_color=INOSYS_COLORS["background_primary"])

        # Center the dialog
        test_dialog.transient(self)
        test_dialog.grab_set()

        # Main frame
        main_frame = ctk.CTkFrame(test_dialog, fg_color=INOSYS_COLORS["background_secondary"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Title
        title_label = ctk.CTkLabel(
            main_frame,
            text="Window Memory Test",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=INOSYS_COLORS["light_blue"]
        )
        title_label.pack(pady=10)

        # Test button
        test_button = ctk.CTkButton(
            main_frame,
            text="Run Test",
            command=self.test_window_memory,
            fg_color=INOSYS_COLORS["medium_blue"],
            hover_color=INOSYS_COLORS["light_blue"]
        )
        test_button.pack(pady=10)

        # Show log button
        log_button = ctk.CTkButton(
            main_frame,
            text="Show Log File",
            command=self.open_window_log,
            fg_color=INOSYS_COLORS["dark_blue"],
            hover_color=INOSYS_COLORS["medium_blue"]
        )
        log_button.pack(pady=5)

        # Close button
        close_button = ctk.CTkButton(
            main_frame,
            text="Close",
            command=test_dialog.destroy,
            fg_color=INOSYS_COLORS["dark_blue"],
            hover_color=INOSYS_COLORS["medium_blue"]
        )
        close_button.pack(pady=10)

    # Also add this method to handle window resize/move events
    def on_window_configure(self, event=None) -> None:
        """Handle window configure events (resize/move)"""
        # Only save if the event is for the main window (not child widgets)
        if event and event.widget == self:
            # Debounce - only save after user stops resizing/moving
            if hasattr(self, '_configure_timer'):
                self.after_cancel(self._configure_timer)

            # Save geometry after 1 second of no changes
            self._configure_timer = self.after(1000, self.save_window_geometry)

    def on_window_interact(self, event=None) -> None:
        """Handle window interaction events"""
        # Save geometry after user interaction (debounced)
        if hasattr(self, '_interact_timer'):
            self.after_cancel(self._interact_timer)

        self._interact_timer = self.after(2000, self.save_window_geometry)

    def save_case_notes(self, notes: str) -> None:
        """Save case notes to current plan"""
        self.current_case_notes = notes

    def load_case_notes(self) -> str:
        """Load case notes from current plan"""
        return getattr(self, 'current_case_notes', '')


if __name__ == "__main__":
    app: PrimusImplantApp = PrimusImplantApp()
    app.mainloop()
