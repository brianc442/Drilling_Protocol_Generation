import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib import colors
import os
import math

# Global variables
implant_data = None
# List to store dictionaries of {tooth_number: str, implant_line_var: StringVar, ...}
# This will hold the current state of all dynamically added implant plans
tooth_plans = []

# Define standard options for dropdowns
IMPLANT_LINE_OPTIONS = ['Primus']
IMPLANT_DIAMETER_OPTIONS = ['3.5', '4.0', '4.5', '5.0']
IMPLANT_LENGTH_OPTIONS = ['7.5', '8.5', '10.0', '11.5', '13.0']
OFFSET_OPTIONS = ['10', '11.5', '13']

# --- Internal CSV Loading ---
# IMPORTANT: Ensure 'Primus Implant List - Primus Implant List.csv' is in the same directory as this script.
INTERNAL_CSV_FILENAME = 'Primus Implant List - Primus Implant List.csv'


def load_internal_csv():
    global implant_data
    try:
        implant_data = pd.read_csv(INTERNAL_CSV_FILENAME, header=0)
        print(f"Internal CSV '{INTERNAL_CSV_FILENAME}' loaded successfully.")
        print("CSV Data Head:\n", implant_data.head())
        print("CSV Columns:\n", implant_data.columns.tolist())
    except FileNotFoundError:
        messagebox.showerror("Error", f"The required CSV file '{INTERNAL_CSV_FILENAME}' was not found.\n"
                                      "Please ensure it is in the same directory as this application.")
        implant_data = None
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load internal CSV: {e}")
        implant_data = None


# --- Drilling Sequence Logic ---
def generate_drilling_sequence(row):
    """
    Extracts the drilling sequence from the provided CSV row.
    It looks for columns like 'Starter Drill', 'Initial Drill 1', 'Drill 1', etc.
    and concatenates their values into a formatted string.
    """
    sequence_parts = []
    drill_columns = [
        'Starter Drill', 'Initial Drill 1', 'Initial Drill 2',
        'Drill 1', 'Drill 2', 'Drill 3', 'Drill 4'
    ]

    for col in drill_columns:
        if col in row and pd.notna(row[col]) and str(row[col]).strip() != '':
            # Append the drill value, converting to string and stripping whitespace
            # Only append "mm" if it's a numeric drill size
            try:
                float(str(row[col]).strip())
                sequence_parts.append(str(row[col]).strip() + "mm")
            except ValueError:
                # If it's not a number, append as is
                sequence_parts.append(str(row[col]).strip())

    if not sequence_parts:
        return "No specific drilling sequence found."

    return " -> ".join(sequence_parts)


# --- PDF Generation ---
def generate_report():
    if implant_data is None:
        messagebox.showwarning("Warning", "Implant data is not loaded. Please ensure the CSV file is present.")
        return

    # Get general report details
    patient_name = patient_entry.get().strip()
    case_id = case_id_entry.get().strip()

    if not all([patient_name, case_id]):
        messagebox.showwarning("Warning", "Please fill in Patient Name and Case ID.")
        return

    if not tooth_plans:
        messagebox.showwarning("Warning", "Please select at least one tooth and configure its implant plan.")
        return

    report_filename = f"Primus_Drill_Report_{patient_name.replace(' ', '_')}_{case_id}.pdf"
    save_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        initialfile=report_filename,
        filetypes=[("PDF files", "*.pdf")]
    )

    if not save_path:
        return  # User cancelled the save dialog

    try:
        doc = SimpleDocTemplate(save_path, pagesize=letter)
        styles = getSampleStyleSheet()
        story = []

        # Add title
        story.append(Paragraph("<b>Primus Dental Implant Drill Report</b>", styles['h1']))
        story.append(Spacer(1, 0.2 * inch))

        # Add patient and case information
        story.append(Paragraph(f"<b>Patient Name:</b> {patient_name}", styles['Normal']))
        story.append(Paragraph(f"<b>Case ID:</b> {case_id}", styles['Normal']))
        story.append(Spacer(1, 0.4 * inch))

        # Define the expected columns from the CSV for display and lookup
        required_data_cols = ['Implant Part No', 'Implant Diameter', 'Implant Length', 'Offset', 'Implant Line']
        drill_sequence_cols = [
            'Starter Drill', 'Initial Drill 1', 'Initial Drill 2',
            'Drill 1', 'Drill 2', 'Drill 3', 'Drill 4'
        ]
        all_required_cols = required_data_cols + drill_sequence_cols

        # Check if all required columns exist in the loaded DataFrame
        missing_cols = [col for col in all_required_cols if col not in implant_data.columns]
        if missing_cols:
            messagebox.showerror(
                "CSV Error",
                f"The loaded CSV is missing one or more required columns: {', '.join(missing_cols)}. "
                "Please ensure the CSV file has the correct headers."
            )
            return

        # Prepare data for the table - for all selected implants
        table_data = [
            ['Tooth No.', 'Implant Part No', 'Diameter (mm)', 'Length (mm)', 'Offset (mm)', 'Drilling Sequence']
        ]

        for plan in tooth_plans:
            tooth_number = plan['tooth_number']
            implant_line = plan['implant_line_var'].get()
            implant_diameter = plan['implant_diameter_var'].get()
            implant_length = plan['implant_length_var'].get()
            offset = plan['offset_var'].get()

            # Convert to float for comparison with CSV data
            try:
                implant_diameter_float = float(implant_diameter)
                implant_length_float = float(implant_length)
                offset_float = float(offset)
            except ValueError:
                messagebox.showerror("Input Error",
                                     f"Invalid numeric input for tooth {tooth_number}. Please check diameter, length, or offset.")
                return

            # Filter the implant_data DataFrame based on user selections for this tooth
            filtered_implant = implant_data[
                (implant_data['Implant Line'] == implant_line) &
                (implant_data['Implant Diameter'] == implant_diameter_float) &
                (implant_data['Implant Length'] == implant_length_float) &
                (implant_data['Offset'] == offset_float)
                ]

            if filtered_implant.empty:
                drilling_seq = "No matching implant found for this selection."
                implant_part_no = "N/A"
            else:
                # Take the first match if multiple exist (shouldn't happen with unique combinations)
                selected_implant_row = filtered_implant.iloc[0]
                implant_part_no = selected_implant_row['Implant Part No']
                drilling_seq = generate_drilling_sequence(selected_implant_row)

            table_data.append([
                str(tooth_number),
                str(implant_part_no),
                str(implant_diameter),
                str(implant_length),
                str(offset),
                drilling_seq
            ])

        # Create the table
        # Adjust colWidths to accommodate the new Tooth No. column
        table = Table(table_data, colWidths=[0.8 * inch, 1.2 * inch, 0.8 * inch, 0.8 * inch, 0.8 * inch, 2.5 * inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#004d40")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#e0f7fa")),
            ('GRID', (0, 0), (-1, -1), 1, colors.HexColor("#00796b")),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('WORDWRAP', (5, 1), (5, -1), True),  # Word wrap for drilling sequence (now column 5)
        ]))
        story.append(table)
        story.append(Spacer(1, 0.5 * inch))
        story.append(Paragraph(
            "<i>Note: This report is generated based on the provided implant list and selected parameters. Always refer to the specific Primus surgical protocol and your clinical judgment.</i>",
            styles['Italic']))

        doc.build(story)
        messagebox.showinfo("Success", f"Drill report saved to:\n{save_path}")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate PDF report: {e}")
        print(f"Error details: {e}")


# --- Tooth Diagram and Dynamic Implant Planning ---

# Dictionary to map tooth numbers to their canvas IDs and current state
tooth_canvas_map = {}
selected_tooth_color = "#a7ffeb"  # Light green for selected teeth
default_tooth_color = "#ffffff"  # White for unselected teeth


def draw_tooth_diagram(canvas):
    """Draws a simplified, numbered 32-tooth diagram on the canvas in two linear rows."""
    canvas.delete("all")  # Clear existing drawings
    tooth_canvas_map.clear()  # Clear previous mappings

    tooth_width = 25  # Adjusted from 30 to 25
    tooth_height = 30
    horizontal_gap = 3  # Adjusted from 5 to 3
    vertical_gap = 10  # Gap between the two rows

    canvas_width = canvas.winfo_width()
    canvas_height = canvas.winfo_height()

    # Calculate total width for a row of 16 teeth
    row_total_width = (16 * tooth_width) + (15 * horizontal_gap)

    # Calculate starting X position to center the rows
    start_x = (canvas_width - row_total_width) / 2

    # Top row (1-16)
    current_x = start_x
    current_y = 30  # A bit from the top edge
    for tooth_num in range(1, 17):
        x1, y1 = current_x, current_y
        x2, y2 = x1 + tooth_width, y1 + tooth_height
        rect_id = canvas.create_rectangle(x1, y1, x2, y2, fill=default_tooth_color, outline="#004d40", width=1)
        text_id = canvas.create_text(x1 + tooth_width / 2, y1 + tooth_height / 2, text=str(tooth_num),
                                     font=("Inter", 9, "bold"), fill="#004d40")
        tooth_canvas_map[str(tooth_num)] = {'rect_id': rect_id, 'text_id': text_id, 'x1': x1, 'y1': y1, 'x2': x2,
                                            'y2': y2, 'selected': False}
        current_x += (tooth_width + horizontal_gap)

    # Bottom row (32-17)
    current_x = start_x
    current_y = current_y + tooth_height + vertical_gap  # Position below the first row
    for tooth_num in range(32, 16, -1):  # Iterate from 32 down to 17
        x1, y1 = current_x, current_y
        x2, y2 = x1 + tooth_width, y1 + tooth_height
        rect_id = canvas.create_rectangle(x1, y1, x2, y2, fill=default_tooth_color, outline="#004d40", width=1)
        text_id = canvas.create_text(x1 + tooth_width / 2, y1 + tooth_height / 2, text=str(tooth_num),
                                     font=("Inter", 9, "bold"), fill="#004d40")
        tooth_canvas_map[str(tooth_num)] = {'rect_id': rect_id, 'text_id': text_id, 'x1': x1, 'y1': y1, 'x2': x2,
                                            'y2': y2, 'selected': False}
        current_x += (tooth_width + horizontal_gap)

    # Re-highlight already selected teeth if any were previously selected
    for plan in tooth_plans:
        tooth_num_str = plan['tooth_number']
        if tooth_num_str in tooth_canvas_map:
            canvas.itemconfig(tooth_canvas_map[tooth_num_str]['rect_id'], fill=selected_tooth_color)
            tooth_canvas_map[tooth_num_str]['selected'] = True


def on_tooth_click(event):
    """Handles clicks on the tooth diagram."""
    canvas = event.widget
    for tooth_num, tooth_info in tooth_canvas_map.items():
        x1, y1, x2, y2 = tooth_info['x1'], tooth_info['y1'], tooth_info['x2'], tooth_info['y2']
        # Check if click is within the bounding box of the tooth
        if x1 <= event.x <= x2 and y1 <= event.y <= y2:
            # Found the clicked tooth
            if tooth_info['selected']:
                # Deselect tooth
                canvas.itemconfig(tooth_info['rect_id'], fill=default_tooth_color)
                tooth_info['selected'] = False
                remove_tooth_plan(tooth_num)
            else:
                # Select tooth
                canvas.itemconfig(tooth_info['rect_id'], fill=selected_tooth_color)
                tooth_info['selected'] = True
                add_tooth_plan(tooth_num)
            break


def add_tooth_plan(tooth_num):
    """Adds a new implant planning section for the selected tooth."""
    # Check if this tooth already has a plan (shouldn't happen with proper logic, but for safety)
    if any(p['tooth_number'] == tooth_num for p in tooth_plans):
        print(f"Warning: Tooth {tooth_num} already has a plan.")
        return

    # Create a new frame for this tooth's implant details
    tooth_frame = ctk.CTkFrame(implant_plans_scroll_frame, fg_color="#f0f9ff", corner_radius=8, border_color="#00796b",
                               border_width=1)
    tooth_frame.pack(fill="x", padx=5, pady=5)
    tooth_frame.grid_columnconfigure(1, weight=1)  # Allow dropdowns to expand

    # Store CTkStringVars to easily retrieve selected values later
    implant_line_var = ctk.StringVar(value=IMPLANT_LINE_OPTIONS[0])
    implant_diameter_var = ctk.StringVar(value=IMPLANT_DIAMETER_OPTIONS[0])
    implant_length_var = ctk.StringVar(value=IMPLANT_LENGTH_OPTIONS[0])
    offset_var = ctk.StringVar(value=OFFSET_OPTIONS[0])

    # Add widgets to the tooth frame
    ctk.CTkLabel(tooth_frame, text=f"Tooth #{tooth_num}", font=ctk.CTkFont(family="Inter", size=13, weight="bold"),
                 text_color="#004d40").grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")

    # Implant Line
    ctk.CTkLabel(tooth_frame, text="Line:", font=ctk.CTkFont(family="Inter", size=11), text_color="#004d40").grid(row=1,
                                                                                                                  column=0,
                                                                                                                  padx=10,
                                                                                                                  pady=2,
                                                                                                                  sticky="w")
    ctk.CTkOptionMenu(tooth_frame, values=IMPLANT_LINE_OPTIONS, variable=implant_line_var,
                      font=ctk.CTkFont(family="Inter", size=11), fg_color="white", text_color="black",
                      button_color="#00796b", button_hover_color="#004d40").grid(row=1, column=1, padx=10, pady=2,
                                                                                 sticky="ew")

    # Implant Diameter
    ctk.CTkLabel(tooth_frame, text="Diameter (mm):", font=ctk.CTkFont(family="Inter", size=11),
                 text_color="#004d40").grid(row=2, column=0, padx=10, pady=2, sticky="w")
    ctk.CTkOptionMenu(tooth_frame, values=IMPLANT_DIAMETER_OPTIONS, variable=implant_diameter_var,
                      font=ctk.CTkFont(family="Inter", size=11), fg_color="white", text_color="black",
                      button_color="#00796b", button_hover_color="#004d40").grid(row=2, column=1, padx=10, pady=2,
                                                                                 sticky="ew")

    # Implant Length
    ctk.CTkLabel(tooth_frame, text="Length (mm):", font=ctk.CTkFont(family="Inter", size=11),
                 text_color="#004d40").grid(row=3, column=0, padx=10, pady=2, sticky="w")
    ctk.CTkOptionMenu(tooth_frame, values=IMPLANT_LENGTH_OPTIONS, variable=implant_length_var,
                      font=ctk.CTkFont(family="Inter", size=11), fg_color="white", text_color="black",
                      button_color="#00796b", button_hover_color="#004d40").grid(row=3, column=1, padx=10, pady=2,
                                                                                 sticky="ew")

    # Offset
    ctk.CTkLabel(tooth_frame, text="Offset (mm):", font=ctk.CTkFont(family="Inter", size=11),
                 text_color="#004d40").grid(row=4, column=0, padx=10, pady=2, sticky="w")
    ctk.CTkOptionMenu(tooth_frame, values=OFFSET_OPTIONS, variable=offset_var,
                      font=ctk.CTkFont(family="Inter", size=11), fg_color="white", text_color="black",
                      button_color="#00796b", button_hover_color="#004d40").grid(row=4, column=1, padx=10, pady=2,
                                                                                 sticky="ew")

    # Remove button for this specific tooth plan
    remove_button = ctk.CTkButton(
        tooth_frame,
        text="Remove",
        command=lambda tn=tooth_num: remove_tooth_plan(tn),
        font=ctk.CTkFont(family="Inter", size=11, weight="bold"),
        fg_color="red",
        hover_color="#cc0000",
        text_color="white",
        corner_radius=5,
        height=25
    )
    remove_button.grid(row=0, column=1, padx=10, pady=(10, 5), sticky="e")  # Place next to tooth number label

    # Store the plan details and references to the frame and variables
    tooth_plans.append({
        'tooth_number': tooth_num,
        'frame': tooth_frame,
        'implant_line_var': implant_line_var,
        'implant_diameter_var': implant_diameter_var,
        'implant_length_var': implant_length_var,
        'offset_var': offset_var
    })
    # Sort tooth_plans by tooth number for consistent display in PDF
    tooth_plans.sort(key=lambda x: int(x['tooth_number']))


def remove_tooth_plan(tooth_num_to_remove):
    """Removes an implant planning section for a deselected tooth."""
    global tooth_plans
    # Find and destroy the corresponding frame
    plan_found = False
    for i, plan in enumerate(tooth_plans):
        if plan['tooth_number'] == tooth_num_to_remove:
            plan['frame'].destroy()  # Destroy the CTkFrame widget
            tooth_plans.pop(i)  # Remove from the list
            plan_found = True
            break

    if plan_found:
        # Update the canvas to reflect deselection
        if str(tooth_num_to_remove) in tooth_canvas_map:
            canvas_info = tooth_canvas_map[str(tooth_num_to_remove)]
            tooth_diagram_canvas.itemconfig(canvas_info['rect_id'], fill=default_tooth_color)
            canvas_info['selected'] = False
    else:
        print(f"Error: Could not find plan for tooth {tooth_num_to_remove} to remove.")


# --- GUI Setup ---
def create_app():
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    root = ctk.CTk()
    root.title("Primus Drill Report Generator")
    root.geometry("800x900")  # Increased size for more content
    root.grid_columnconfigure(0, weight=1)  # Allow column to expand
    root.grid_rowconfigure(3, weight=1)  # Allow the scrollable frame to expand

    # Load the internal CSV data when the application starts
    load_internal_csv()

    # Title Label
    title_label = ctk.CTkLabel(
        root,
        text="Primus Dental Implant Drill Report Generator",
        font=ctk.CTkFont(family="Inter", size=20, weight="bold"),
        text_color="#004d40"
    )
    title_label.grid(row=0, column=0, pady=20, padx=20, sticky="ew")

    # Patient & Case ID Input Section
    general_input_frame = ctk.CTkFrame(root, fg_color="#e0f7fa", corner_radius=10, border_color="#004d40",
                                       border_width=2)
    general_input_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
    general_input_frame.grid_columnconfigure(0, weight=1)
    general_input_frame.grid_columnconfigure(1, weight=3)

    general_input_frame_title = ctk.CTkLabel(general_input_frame, text="1. Enter General Report Details",
                                             font=ctk.CTkFont(family="Inter", size=14, weight="bold"),
                                             text_color="#004d40")
    general_input_frame_title.grid(row=0, column=0, columnspan=2, padx=10, pady=(10, 5), sticky="w")

    # Patient Name
    ctk.CTkLabel(general_input_frame, text="Patient Name:", font=ctk.CTkFont(family="Inter", size=11),
                 text_color="#004d40").grid(row=1, column=0, padx=10, pady=5, sticky="w")
    global patient_entry
    patient_entry = ctk.CTkEntry(general_input_frame, font=ctk.CTkFont(family="Inter", size=11), width=40,
                                 corner_radius=5, border_width=1, fg_color="white", text_color="black")
    patient_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

    # Case ID
    ctk.CTkLabel(general_input_frame, text="Case ID:", font=ctk.CTkFont(family="Inter", size=11),
                 text_color="#004d40").grid(row=2, column=0, padx=10, pady=5, sticky="w")
    global case_id_entry
    case_id_entry = ctk.CTkEntry(general_input_frame, font=ctk.CTkFont(family="Inter", size=11), width=40,
                                 corner_radius=5, border_width=1, fg_color="white", text_color="black")
    case_id_entry.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

    # Tooth Selection & Implant Planning Section
    implant_planning_frame = ctk.CTkFrame(root, fg_color="#e0f7fa", corner_radius=10, border_color="#004d40",
                                          border_width=2)
    implant_planning_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
    implant_planning_frame.grid_columnconfigure(0, weight=1)

    implant_planning_frame_title = ctk.CTkLabel(implant_planning_frame, text="2. Select Teeth & Plan Implants",
                                                font=ctk.CTkFont(family="Inter", size=14, weight="bold"),
                                                text_color="#004d40")
    implant_planning_frame_title.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="w")

    # Tooth Diagram Canvas
    global tooth_diagram_canvas  # Make global for drawing and event binding
    tooth_diagram_canvas = ctk.CTkCanvas(implant_planning_frame, width=550, height=100, bg="white",
                                         highlightbackground="#004d40", highlightthickness=2)  # Adjusted height
    tooth_diagram_canvas.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
    tooth_diagram_canvas.bind("<Button-1>", on_tooth_click)

    # Draw the initial diagram
    # Use root.update_idletasks() to ensure canvas has determined its size before drawing
    root.update_idletasks()
    draw_tooth_diagram(tooth_diagram_canvas)

    # Frame to hold dynamically added implant plan details (scrollable)
    global implant_plans_scroll_frame  # Make global to add frames to it
    implant_plans_scroll_frame = ctk.CTkScrollableFrame(root, fg_color="#e0f7fa", corner_radius=10,
                                                        border_color="#004d40", border_width=2)
    implant_plans_scroll_frame.grid(row=3, column=0, padx=20, pady=10, sticky="nsew")  # This frame expands vertically
    implant_plans_scroll_frame.grid_columnconfigure(0, weight=1)  # Allow content inside to expand

    # Generate Report Button
    generate_button = ctk.CTkButton(
        root,
        text="Generate Drill Report PDF",
        command=generate_report,
        font=ctk.CTkFont(family="Inter", size=14, weight="bold"),
        fg_color="#004d40",
        hover_color="#00796b",
        text_color="white",
        corner_radius=10,
        height=50
    )
    generate_button.grid(row=4, column=0, pady=20, padx=20, sticky="ew")

    root.mainloop()


if __name__ == "__main__":
    create_app()
