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
from datetime import datetime
from typing import Dict, List, Any, Optional, Callable
from PIL import Image as PILImage

# Set appearance mode and color theme
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# Pantone Color Definitions
# Pantone 3597C - Light Blue
PANTONE_3597C = "#00B5D8"  # Light blue
# Pantone 7683C - Dark Blue
PANTONE_7683C = "#1E3A8A"  # Dark blue
# Pantone 659C - Medium Blue
PANTONE_659C = "#0EA5E9"  # Medium blue

# Color scheme dictionary for easy reference
INOSYS_COLORS = {
    "light_blue": PANTONE_3597C,
    "dark_blue": PANTONE_7683C,
    "medium_blue": PANTONE_659C,
    "white": "#FFFFFF",
    "light_gray": "#F8F9FA",
    "dark_gray": "#6B7280"
}


class ToothDiagram(ctk.CTkFrame):
    def __init__(self, parent: ctk.CTkFrame, callback: Callable[[int], None]) -> None:
        super().__init__(parent)
        self.callback: Callable[[int], None] = callback
        self.selected_tooth: Optional[int] = None
        self.tooth_buttons: Dict[int, ctk.CTkButton] = {}

        # Tooth numbering (Universal Numeric Notation)
        # Upper teeth: 1-16 (right to left)
        # Lower teeth: 17-32 (left to right)

        self.create_tooth_diagram()

    def create_tooth_diagram(self) -> None:
        title: ctk.CTkLabel = ctk.CTkLabel(
            self,
            text="Select Tooth (Universal Numeric Notation)",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=INOSYS_COLORS["dark_blue"]
        )
        title.pack(pady=10)

        # Upper teeth frame
        upper_frame: ctk.CTkFrame = ctk.CTkFrame(self, fg_color=INOSYS_COLORS["light_gray"])
        upper_frame.pack(pady=5)

        upper_label: ctk.CTkLabel = ctk.CTkLabel(
            upper_frame,
            text="Upper Teeth",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=INOSYS_COLORS["dark_blue"]
        )
        upper_label.pack(pady=5)

        upper_teeth_frame: ctk.CTkFrame = ctk.CTkFrame(upper_frame, fg_color=INOSYS_COLORS["white"])
        upper_teeth_frame.pack(pady=5)

        # Upper teeth (1-16, right to left)
        for i in range(1, 17):
            btn: ctk.CTkButton = ctk.CTkButton(
                upper_teeth_frame,
                text=str(i),
                width=40,
                height=40,
                command=lambda tooth=i: self.select_tooth(tooth),
                fg_color=INOSYS_COLORS["medium_blue"],
                hover_color=INOSYS_COLORS["light_blue"],
                text_color=INOSYS_COLORS["white"],
                font=ctk.CTkFont(size=12, weight="bold")
            )
            btn.grid(row=0, column=16 - i, padx=2, pady=2)
            self.tooth_buttons[i] = btn

        # Lower teeth frame
        lower_frame: ctk.CTkFrame = ctk.CTkFrame(self, fg_color=INOSYS_COLORS["light_gray"])
        lower_frame.pack(pady=5)

        lower_label: ctk.CTkLabel = ctk.CTkLabel(
            lower_frame,
            text="Lower Teeth",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color=INOSYS_COLORS["dark_blue"]
        )
        lower_label.pack(pady=5)

        lower_teeth_frame: ctk.CTkFrame = ctk.CTkFrame(lower_frame, fg_color=INOSYS_COLORS["white"])
        lower_teeth_frame.pack(pady=5)

        # Lower teeth (17-32, left to right)
        for i in range(17, 33):
            btn: ctk.CTkButton = ctk.CTkButton(
                lower_teeth_frame,
                text=str(i),
                width=40,
                height=40,
                command=lambda tooth=i: self.select_tooth(tooth),
                fg_color=INOSYS_COLORS["medium_blue"],
                hover_color=INOSYS_COLORS["light_blue"],
                text_color=INOSYS_COLORS["white"],
                font=ctk.CTkFont(size=12, weight="bold")
            )
            btn.grid(row=0, column=i - 17, padx=2, pady=2)
            self.tooth_buttons[i] = btn

    def select_tooth(self, tooth_num: int) -> None:
        # Reset all buttons to default color
        for btn in self.tooth_buttons.values():
            btn.configure(fg_color=INOSYS_COLORS["medium_blue"])

        # Highlight selected tooth
        self.tooth_buttons[tooth_num].configure(fg_color=INOSYS_COLORS["dark_blue"])
        self.selected_tooth = tooth_num
        self.callback(tooth_num)


class PrimusImplantApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()

        self.title("Primus Implant Report Generator")
        self.geometry("1000x800")

        # Configure window colors
        self.configure(fg_color=INOSYS_COLORS["light_gray"])

        # Initialize instance variables with type hints
        self.implant_data: pd.DataFrame = pd.DataFrame()
        self.implant_plans: List[Dict[str, Any]] = []

        # GUI components - will be initialized in create_widgets
        self.notebook: ctk.CTkTabview
        self.tooth_diagram: ToothDiagram
        self.selected_tooth_label: ctk.CTkLabel
        self.implant_line_var: ctk.StringVar
        self.implant_line_combo: ctk.CTkComboBox
        self.implant_diameter_var: ctk.StringVar
        self.implant_diameter_combo: ctk.CTkComboBox
        self.implant_length_var: ctk.StringVar
        self.implant_length_combo: ctk.CTkComboBox
        self.offset_var: ctk.StringVar
        self.offset_combo: ctk.CTkComboBox
        self.plan_scrollable_frame: ctk.CTkScrollableFrame
        self.doctor_name_entry: ctk.CTkEntry
        self.patient_name_entry: ctk.CTkEntry
        self.case_number_entry: ctk.CTkEntry

        # Load CSV data
        self.load_implant_data()

        self.create_widgets()

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
        main_frame: ctk.CTkFrame = ctk.CTkFrame(self, fg_color=INOSYS_COLORS["white"])
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Header frame for logo and title
        header_frame: ctk.CTkFrame = ctk.CTkFrame(main_frame, fg_color=INOSYS_COLORS["dark_blue"])
        header_frame.pack(fill="x", pady=(0, 20))

        # Load and display logo
        self.load_and_display_logo(header_frame)

        # Title
        title_label: ctk.CTkLabel = ctk.CTkLabel(
            header_frame,
            text="Primus Implant Report Generator",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color=INOSYS_COLORS["white"]
        )
        title_label.pack(pady=10)

        # Create notebook for tabs
        self.notebook = ctk.CTkTabview(
            main_frame,
            fg_color=INOSYS_COLORS["light_gray"],
            segmented_button_fg_color=INOSYS_COLORS["medium_blue"],
            segmented_button_selected_color=INOSYS_COLORS["dark_blue"],
            segmented_button_selected_hover_color=INOSYS_COLORS["dark_blue"]
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

    def setup_add_implant_tab(self) -> None:
        tab: ctk.CTkFrame = self.notebook.tab("Add Implant")

        # Create scrollable frame
        scrollable_frame: ctk.CTkScrollableFrame = ctk.CTkScrollableFrame(tab)
        scrollable_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Tooth selection
        tooth_frame: ctk.CTkFrame = ctk.CTkFrame(scrollable_frame)
        tooth_frame.pack(fill="x", padx=10, pady=10)

        self.tooth_diagram = ToothDiagram(tooth_frame, self.on_tooth_selected)
        self.tooth_diagram.pack(fill="both", expand=True, padx=10, pady=10)

        # Selected tooth display
        self.selected_tooth_label = ctk.CTkLabel(
            scrollable_frame,
            text="Selected Tooth: None",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        self.selected_tooth_label.pack(pady=10)

        # Input fields frame
        input_frame: ctk.CTkFrame = ctk.CTkFrame(scrollable_frame)
        input_frame.pack(fill="x", padx=10, pady=10)

        # Implant Line
        ctk.CTkLabel(
            input_frame,
            text="Implant Line:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.implant_line_var = ctk.StringVar(value="Primus")
        self.implant_line_combo = ctk.CTkComboBox(
            input_frame,
            values=["Primus"],
            variable=self.implant_line_var,
            state="readonly"
        )
        self.implant_line_combo.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        # Implant Diameter
        ctk.CTkLabel(
            input_frame,
            text="Implant Diameter:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).grid(row=1, column=0, padx=10, pady=5, sticky="w")

        self.implant_diameter_var = ctk.StringVar()
        self.implant_diameter_combo = ctk.CTkComboBox(
            input_frame,
            values=["3.5", "4.0", "4.5", "5.0"],
            variable=self.implant_diameter_var
        )
        self.implant_diameter_combo.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        # Implant Length
        ctk.CTkLabel(
            input_frame,
            text="Implant Length:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).grid(row=2, column=0, padx=10, pady=5, sticky="w")

        self.implant_length_var = ctk.StringVar()
        self.implant_length_combo = ctk.CTkComboBox(
            input_frame,
            values=["7.5", "8.5", "10.0", "11.5", "13.0"],
            variable=self.implant_length_var
        )
        self.implant_length_combo.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        # Offset
        ctk.CTkLabel(
            input_frame,
            text="Offset:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).grid(row=3, column=0, padx=10, pady=5, sticky="w")

        self.offset_var = ctk.StringVar()
        self.offset_combo = ctk.CTkComboBox(
            input_frame,
            values=["10", "11.5", "13"],
            variable=self.offset_var
        )
        self.offset_combo.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        # Configure grid weights
        input_frame.columnconfigure(1, weight=1)

        # Add implant button
        add_button: ctk.CTkButton = ctk.CTkButton(
            scrollable_frame,
            text="Add Implant to Plan",
            command=self.add_implant_to_plan,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=INOSYS_COLORS["dark_blue"],
            hover_color=INOSYS_COLORS["medium_blue"],
            text_color=INOSYS_COLORS["white"]
        )
        add_button.pack(pady=20)

    def setup_review_plan_tab(self) -> None:
        tab: ctk.CTkFrame = self.notebook.tab("Review Plan")

        # Frame for the plan list
        plan_frame: ctk.CTkFrame = ctk.CTkFrame(tab)
        plan_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Title
        title_label: ctk.CTkLabel = ctk.CTkLabel(
            plan_frame,
            text="Current Implant Plan",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title_label.pack(pady=10)

        # Scrollable frame for plan items
        self.plan_scrollable_frame = ctk.CTkScrollableFrame(plan_frame)
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
        report_frame: ctk.CTkFrame = ctk.CTkFrame(tab)
        report_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Title
        title_label: ctk.CTkLabel = ctk.CTkLabel(
            report_frame,
            text="Generate PDF Report",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        title_label.pack(pady=20)

        # Doctor information
        info_frame: ctk.CTkFrame = ctk.CTkFrame(report_frame)
        info_frame.pack(fill="x", padx=20, pady=10)

        ctk.CTkLabel(
            info_frame,
            text="Doctor Name:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.doctor_name_entry = ctk.CTkEntry(info_frame, width=300)
        self.doctor_name_entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(
            info_frame,
            text="Patient Name:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).grid(row=1, column=0, padx=10, pady=5, sticky="w")

        self.patient_name_entry = ctk.CTkEntry(info_frame, width=300)
        self.patient_name_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        ctk.CTkLabel(
            info_frame,
            text="Case Number:",
            font=ctk.CTkFont(size=12, weight="bold")
        ).grid(row=2, column=0, padx=10, pady=5, sticky="w")

        self.case_number_entry = ctk.CTkEntry(info_frame, width=300)
        self.case_number_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        info_frame.columnconfigure(1, weight=1)

        # Generate button
        generate_button: ctk.CTkButton = ctk.CTkButton(
            report_frame,
            text="Generate PDF Report",
            command=self.generate_pdf_report,
            height=50,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color=INOSYS_COLORS["dark_blue"],
            hover_color=INOSYS_COLORS["medium_blue"],
            text_color=INOSYS_COLORS["white"]
        )
        generate_button.pack(pady=30)

    def on_tooth_selected(self, tooth_num: int) -> None:
        self.selected_tooth_label.configure(text=f"Selected Tooth: {tooth_num}")

    def add_implant_to_plan(self) -> None:
        # Validate inputs
        if not self.tooth_diagram.selected_tooth:
            messagebox.showerror("Error", "Please select a tooth first!")
            return

        if not all([self.implant_diameter_var.get(), self.implant_length_var.get(), self.offset_var.get()]):
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

        # Create implant plan
        implant_plan: Dict[str, Any] = {
            'tooth_number': self.tooth_diagram.selected_tooth,
            'implant_line': self.implant_line_var.get(),
            'diameter': diameter,
            'length': length,
            'offset': offset,
            'implant_data': matching_implant.iloc[0].to_dict()
        }

        # Check if tooth already has an implant planned
        existing_plan: Optional[Dict[str, Any]] = next(
            (plan for plan in self.implant_plans if plan['tooth_number'] == implant_plan['tooth_number']),
            None
        )

        if existing_plan:
            if messagebox.askyesno(
                    "Tooth Already Planned",
                    f"Tooth {implant_plan['tooth_number']} already has an implant planned. Replace it?"
            ):
                self.implant_plans.remove(existing_plan)
            else:
                return

        self.implant_plans.append(implant_plan)
        self.update_plan_display()
        messagebox.showinfo("Success", f"Implant added to plan for tooth {implant_plan['tooth_number']}!")

        # Clear selections
        self.implant_diameter_var.set("")
        self.implant_length_var.set("")
        self.offset_var.set("")

    def update_plan_display(self) -> None:
        # Clear existing plan display
        for widget in self.plan_scrollable_frame.winfo_children():
            widget.destroy()

        # Add each implant plan to display
        for i, plan in enumerate(self.implant_plans):
            plan_frame: ctk.CTkFrame = ctk.CTkFrame(self.plan_scrollable_frame)
            plan_frame.pack(fill="x", padx=10, pady=5)

            # Plan details
            details_text: str = (
                f"Tooth {plan['tooth_number']}: {plan['implant_line']} - "
                f"{plan['diameter']}mm × {plan['length']}mm (Offset: {plan['offset']}mm)"
            )
            details_label: ctk.CTkLabel = ctk.CTkLabel(
                plan_frame,
                text=details_text,
                font=ctk.CTkFont(size=12, weight="bold")
            )
            details_label.pack(side="left", padx=10, pady=10)

            # Remove button
            remove_button: ctk.CTkButton = ctk.CTkButton(
                plan_frame,
                text="Remove",
                width=80,
                command=lambda idx=i: self.remove_implant_plan(idx),
                fg_color=INOSYS_COLORS["medium_blue"],
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
            bottomMargin=0.5 * inch
        )
        styles = getSampleStyleSheet()
        story: List[Any] = []

        # Custom styles
        title_style: ParagraphStyle = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=20,
            textColor=colors.Color(30 / 255, 58 / 255, 138 / 255),  # Pantone 7683C - Dark Blue
            alignment=TA_CENTER,
            spaceAfter=30
        )

        header_style: ParagraphStyle = ParagraphStyle(
            'CustomHeader',
            parent=styles['Heading2'],
            fontSize=14,
            textColor=colors.Color(30 / 255, 58 / 255, 138 / 255),  # Pantone 7683C - Dark Blue
            spaceBefore=20,
            spaceAfter=10
        )

        # Add logo to the report if it exists
        logo_added: bool = self.add_logo_to_report(story)

        if logo_added:
            story.append(Spacer(1, 10))

        # Title
        story.append(Paragraph("PRIMUS DENTAL IMPLANT", title_style))
        story.append(Paragraph("SURGICAL DRILLING PROTOCOL", title_style))
        story.append(Spacer(1, 20))

        # Case information
        case_info: List[List[str]] = [
            ["Doctor:", doctor_name],
            ["Patient:", patient_name],
            ["Case Number:", case_number],
            ["Date:", datetime.now().strftime("%B %d, %Y")],
            ["Time:", datetime.now().strftime("%I:%M %p")]
        ]

        case_table: Table = Table(case_info, colWidths=[1.5 * inch, 4 * inch])
        case_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 11),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey)
        ]))

        story.append(case_table)
        story.append(Spacer(1, 30))

        # Sort implant plans by tooth number
        sorted_plans: List[Dict[str, Any]] = sorted(self.implant_plans, key=lambda x: x['tooth_number'])

        # Generate drilling protocol for each implant
        for i, plan in enumerate(sorted_plans):
            # Implant header
            story.append(Paragraph(f"IMPLANT #{i + 1} - TOOTH {plan['tooth_number']}", header_style))

            # Implant specifications
            implant_specs: List[List[str]] = [
                ["Implant Line:", plan['implant_line']],
                ["Part Number:", plan['implant_data']['Implant Part No']],
                ["Diameter:", f"{plan['diameter']} mm"],
                ["Length:", f"{plan['length']} mm"],
                ["Offset:", f"{plan['offset']} mm"],
                ["Guide Sleeve:", plan['implant_data']['Guide Sleeve']],
                ["Drill Length:", f"{plan['implant_data']['Drill Length']} mm"]
            ]

            specs_table: Table = Table(implant_specs, colWidths=[1.5 * inch, 2.5 * inch])
            specs_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                ('FONTNAME', (1, 0), (1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey),
                ('BACKGROUND', (0, 0), (0, -1), colors.Color(0 / 255, 181 / 255, 216 / 255))
                # Pantone 3597C - Light Blue
            ]))

            story.append(specs_table)
            story.append(Spacer(1, 15))

            # Drilling sequence
            story.append(Paragraph(
                "DRILLING SEQUENCE:",
                ParagraphStyle('DrillHeader', parent=styles['Heading3'],
                               fontSize=12, textColor=colors.Color(14 / 255, 165 / 255, 233 / 255))
                # Pantone 659C - Medium Blue
            ))

            # Create drilling sequence table
            drill_data: List[List[str]] = [["Step", "Drill Type", "Depth (mm)", "Notes"]]

            # Add drilling steps
            drill_steps: List[tuple] = [
                ("1", "Starter Drill", plan['implant_data']['Starter Drill'], "Initial pilot hole"),
                ("2", "Initial Drill 1", plan['implant_data']['Initial Drill 1'], "First expansion"),
                ("3", "Initial Drill 2", plan['implant_data']['Initial Drill 2'], "Second expansion"),
                ("4", "Drill 1", plan['implant_data']['Drill 1'], "Progressive drilling"),
                ("5", "Drill 2", plan['implant_data']['Drill 2'], "Progressive drilling"),
                ("6", "Drill 3", plan['implant_data']['Drill 3'], "Progressive drilling"),
                ("7", "Drill 4", plan['implant_data']['Drill 4'], "Final preparation")
            ]

            drill_data.extend(drill_steps)

            drill_table: Table = Table(drill_data, colWidths=[0.7 * inch, 1.2 * inch, 1 * inch, 2.1 * inch])
            drill_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ('BACKGROUND', (0, 0), (-1, 0), colors.Color(30 / 255, 58 / 255, 138 / 255)),
                # Pantone 7683C - Dark Blue
                ('BACKGROUND', (0, 1), (0, -1), colors.Color(14 / 255, 165 / 255, 233 / 255)),
                # Pantone 659C - Medium Blue
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white)
            ]))

            story.append(drill_table)
            story.append(Spacer(1, 20))

            # Important notes for each implant
            notes_text: str = f"""
            <b>IMPORTANT NOTES FOR TOOTH {plan['tooth_number']}:</b><br/>
            • Use copious irrigation during drilling<br/>
            • Maintain drilling speed as per manufacturer guidelines<br/>
            • Check depth frequently with depth gauge<br/>
            • Ensure proper angulation with guide sleeve {plan['implant_data']['Guide Sleeve']}<br/>
            • Final drill depth: {plan['implant_data']['Drill Length']} mm<br/>
            • Implant placement depth should account for {plan['offset']} mm offset
            """

            story.append(Paragraph(notes_text, styles['Normal']))
            story.append(Spacer(1, 20))

        # General surgical notes
        story.append(Paragraph("GENERAL SURGICAL PROTOCOL", header_style))

        general_notes: str = """
        <b>Pre-Surgical Preparation:</b><br/>
        • Verify patient identity and surgical site<br/>
        • Confirm implant specifications and drilling sequence<br/>
        • Prepare sterile surgical field<br/>
        • Check all instruments and drill bits<br/><br/>

        <b>Drilling Protocol:</b><br/>
        • Use intermittent drilling with 15-30 second intervals<br/>
        • Maintain drilling speed: 800-1200 RPM for initial drills, 600-800 RPM for final drills<br/>
        • Apply light pressure - let the drill do the work<br/>
        • Use copious external irrigation (minimum 50ml/min)<br/>
        • Check depth with depth gauge after each drill<br/><br/>

        <b>Post-Drilling:</b><br/>
        • Irrigate the osteotomy thoroughly<br/>
        • Check final depth and angulation<br/>
        • Verify osteotomy diameter with sizing gauge<br/>
        • Proceed with implant placement as per protocol
        """

        story.append(Paragraph(general_notes, styles['Normal']))
        story.append(Spacer(1, 20))

        # Footer
        footer_text: str = f"""
        <b>Report Generated:</b> {datetime.now().strftime("%B %d, %Y at %I:%M %p")}<br/>
        <b>Software:</b> Primus Implant Report Generator v1.0<br/>
        <b>Total Implants:</b> {len(self.implant_plans)}
        """

        story.append(Paragraph(
            footer_text,
            ParagraphStyle('Footer', parent=styles['Normal'],
                           fontSize=8, textColor=colors.grey)
        ))

        # Build PDF
        doc.build(story)

    def add_logo_to_report(self, story: List[Any]) -> bool:
        """Add the Inosys logo to the PDF report"""
        logo_files: List[str] = [
            "inosys_logo.png",
            "inosys_logo.jpg",
            "inosys_logo.jpeg",
            "logo.png",
            "logo.jpg",
            "logo.jpeg"
        ]

        for logo_file in logo_files:
            if os.path.exists(logo_file):
                try:
                    # Create logo image for PDF
                    logo_image = Image(logo_file, width=2 * inch, height=0.8 * inch)
                    logo_image.hAlign = 'LEFT'
                    story.append(logo_image)
                    print(f"Logo added to PDF report from {logo_file}")
                    return True
                except Exception as e:
                    print(f"Error adding logo to PDF from {logo_file}: {str(e)}")
                    continue

        print("Logo file not found for PDF report")
        return False


if __name__ == "__main__":
    app: PrimusImplantApp = PrimusImplantApp()
    app.mainloop()