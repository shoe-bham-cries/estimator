import tkinter as tk
from tkinter import ttk, messagebox
import re
import datetime
import openpyxl
import os
import shutil
import glob
import sys
import pandas as pd
from dotenv import load_dotenv
load_dotenv()  # Load environment variables from .env file

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))

    if relative_path is not None:
        relative_path = os.path.normpath(relative_path)  # Normalize the path separators
        return os.path.join(base_path, relative_path)
    else:
        return None  # or handle the case when relative_path is None



###DEPENDENCIES
csv_path = resource_path(os.getenv('RESOURCE_CSV_PATH'))
if csv_path is not None:
    df = pd.read_csv(csv_path)
else:
    print("CSV file path is not specified.")
    exit(1)
AVAILABLE_PAPER_ROLLS = sorted(df.iloc[:, 1].astype(float).dropna().tolist())
AVAILABLE_CYLINDERS = sorted(df.iloc[:, 0].astype(float).dropna().tolist())
save_directory_env = os.getenv('SAVE_DIRECTORY')
backup_directory_env = os.getenv('BACKUP_DIRECTORY')
temp_path_env = os.getenv('TEMPLATE_PATH')


class RawMaterialEstimatorGUI:
    """
    Graphical User Interface for a Raw Material Estimator.

    This GUI application allows users to estimate raw material requirements for a printing job, such as job details, 
    customer information, dimensions, quantity, printing options, and color details. It calculates and generates a 
    Bill of Material (BOM) based on the provided information.

    The GUI provides input fields, dropdown menus, and buttons for users to enter the necessary data. It performs validation on the 
    input fields to ensure data integrity. The BOM is generated in an Excel spreadsheet, and a copy of the BOM is saved to a specified 
    directory for record-keeping.

    The GUI also incorporates dynamic fields for specifying colors based on the selected printing options. It adjusts the layout 
    dynamically based on the number of color fields required.

    The class methods include event handlers, data validation functions, and utility functions for generating job numbers and handling 
    GUI field updates.

    To use the RawMaterialEstimatorGUI, create an instance of the class and call the `run()` method.

    Attributes:
        window (Tk): The main window of the GUI.
        job_name_entry (Entry): Entry field for job name.
        customer_name_entry (Entry): Entry field for customer name.
        customer_email_entry (Entry): Entry field for customer email.
        customer_mobile_entry (Entry): Entry field for customer mobile number.
        job_type_combo (Combobox): Dropdown menu for selecting job type.
        width_entry (Entry): Entry field for width dimension.
        bottom_entry (Entry): Entry field for bottom dimension.
        height_entry (Entry): Entry field for height dimension.
        gsm_entry (Entry): Entry field for GSM value.
        quantity_entry (Entry): Entry field for quantity.
        printing_var (IntVar): Variable for printing option selection.
        color_entries (list): List of color entry fields.
        color_labels (list): List of color labels.
        submit_button (Button): Button for submitting the data.
        job_number_label (Label): Label for displaying the generated job number.

    Methods:
        __init__(self): Initializes the RawMaterialEstimatorGUI class.
        generate_job_number(self): Generates a unique job number based on the current date and existing files.
        create_widgets(self): Creates and configures the GUI widgets.
        dynamic_fields(self, event): Handles the dynamic creation and removal of color fields based on the printing option.
        validate_fields(self): Validates the input fields to ensure data integrity.
        process_data(self): Processes the validated data and generates the Bill of Material (BOM).
        handle_fields(self): Resets and updates the GUI fields after processing the data.
        run(self): Runs the GUI application by starting the main event loop.

    Note:
    - This script uses the tkinter library for the GUI components.
    - The script assumes the availability of a template.xlsx file in the specified path.

    Author: Shubham R
    Date: 17/05/2023
    """
    def __init__(self):
        """
        Initializes the RawMaterialEstimatorGUI class.

        Creates a GUI window for raw material estimation.
        Sets up the layout and input fields for job details.
        """
        self.window = tk.Tk()
        self.window.title("Raw Material Estimator")

        self.style = ttk.Style()
        self.style.theme_use("clam")

        self.main_frame = ttk.Frame(self.window, padding=20)
        self.main_frame.grid(row=0, column=0, sticky="nsew")

        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)

        # Display the job number at the top
        self.job_number_label = ttk.Label(self.main_frame, text=f"Job Number: {self.generate_job_number()}", font=("Arial", 14, "bold"),
                                           foreground="red")
        self.job_number_label.grid(row=0, column=0, columnspan=2, pady=10)

        # Customer Name entry field
        ttk.Label(self.main_frame, text="Customer Name:", font=("Arial", 12)).grid(row=1, column=0, pady=5, sticky="w")
        self.customer_name_entry = ttk.Entry(self.main_frame, width=45)
        self.customer_name_entry.grid(row=1, column=1, padx=10)

        # Customer Email entry field
        ttk.Label(self.main_frame, text="Customer Email:", font=("Arial", 12)).grid(row=2, column=0, pady=5, sticky="w")
        self.customer_email_entry = ttk.Entry(self.main_frame, width=45)
        self.customer_email_entry.grid(row=2, column=1, padx=10)

        # Job Name entry field
        ttk.Label(self.main_frame, text="Job Name:", font=("Arial", 12)).grid(row=3, column=0, pady=5, sticky="w")
        self.job_name_entry = ttk.Entry(self.main_frame, width=45)
        self.job_name_entry.grid(row=3, column=1, padx=10)

        # Customer Mobile entry field
        ttk.Label(self.main_frame, text="Customer Mobile:", font=("Arial", 12)).grid(row=4, column=0, pady=5, sticky="w")
        self.customer_mobile_entry = ttk.Entry(self.main_frame, width=30)
        self.customer_mobile_entry.grid(row=4, column=1, padx=10)

        # Job Type dropdown
        ttk.Label(self.main_frame, text="Job Type:", font=("Arial", 12)).grid(row=5, column=0, pady=5, sticky="w")
        self.job_type_combo = ttk.Combobox(self.main_frame, values=["SOS", "Carry Bag", "V-Bottom", "Thumb Cut", "Square Cut"], 
                                           font=("Arial", 12))
        self.job_type_combo.current(0)
        self.job_type_combo.grid(row=5, column=1, padx=10)

        # Width entry field
        ttk.Label(self.main_frame, text="Width (in):", font=("Arial", 12)).grid(row=6, column=0, pady=5, sticky="w")
        self.width_entry = ttk.Entry(self.main_frame, width=30)
        self.width_entry.grid(row=6, column=1, padx=10)

        # Bottom entry field
        ttk.Label(self.main_frame, text="Bottom (in):", font=("Arial", 12)).grid(row=7, column=0, pady=5, sticky="w")
        self.bottom_entry = ttk.Entry(self.main_frame, width=30)
        self.bottom_entry.grid(row=7, column=1, padx=10)

        # Height entry field
        ttk.Label(self.main_frame, text="Height (in):", font=("Arial", 12)).grid(row=8, column=0, pady=5, sticky="w")
        self.height_entry = ttk.Entry(self.main_frame, width=30)
        self.height_entry.grid(row=8, column=1, padx=10)

        # GSM entry field
        ttk.Label(self.main_frame, text="GSM:", font=("Arial", 12)).grid(row=9, column=0, pady=5, sticky="w")
        self.gsm_entry = ttk.Entry(self.main_frame, width=30)
        self.gsm_entry.grid(row=9, column=1, padx=10)

        # Quantity entry field
        ttk.Label(self.main_frame, text="Quantity:", font=("Arial", 12)).grid(row=10, column=0, pady=5, sticky="w")
        self.quantity_entry = ttk.Entry(self.main_frame, width=30)
        self.quantity_entry.grid(row=10, column=1, padx=10)

        # Printing dropdown
        ttk.Label(self.main_frame, text="Printing: (0 means no colors)", font=("Arial", 12)).grid(row=11, column=0, pady=5, sticky="w")
        self.printing_var = tk.StringVar()
        self.printing_dropdown = ttk.Combobox(self.main_frame, textvariable=self.printing_var, 
                                              values=["0", "1", "2", "3", "4", "5", "6"], font=("Arial", 12))
        self.printing_dropdown.current(0)
        self.printing_dropdown.grid(row=11, column=1, padx=10)
        self.printing_dropdown.bind("<<ComboboxSelected>>", self.dynamic_fields)

        # Line break between Color Details and Color 1, Color 2, Color 3
        ttk.Label(self.main_frame, text="Color Details:", font=("Arial", 12)).grid(row=12, column=0, pady=5, sticky="e")
        ttk.Label(self.main_frame, text="").grid(row=13, column=0)

        self.color_entries = []  # List to store color entry fields
        self.color_labels = []   # List to store color labels
        self.submit_button = ttk.Button(self.main_frame, text="Submit", command=self.process_data)
        self.submit_button.grid(row=14, column=0, columnspan=2, pady=10, sticky="nsew")

    def dynamic_fields(self, event):
        """
        Update the color fields based on the selected printing option.

        Args:
            event (Event): The event object triggered by the printing option selection.

        Comments:
        - Destroys existing color entries and labels.
        - Creates color labels and entries based on the selected printing option.
        - Removes excess color fields if the printing option is reduced.
        - Configures the submit button and adjusts the grid layout accordingly.
        - Updates the window to reflect the changes.

        """

        # Destroy existing color entries and labels
        for entry in self.color_entries:
            entry.destroy()
        self.color_entries.clear()

        for label in self.color_labels:
            label.destroy()
        self.color_labels.clear()

        printing_value = int(self.printing_var.get())

        # Create color labels and entries
        for i in range(printing_value):
            label = ttk.Label(self.main_frame, text=f"Color {i+1}:", font=("Arial", 12))
            label.grid(row=i+13, column=0, pady=5, sticky="e")
            self.color_labels.append(label)

            color_entry = ttk.Entry(self.main_frame, width=30)
            color_entry.grid(row=i+13, column=1, padx=10, sticky="w")
            self.color_entries.append(color_entry)

        # Remove excess color fields if the printing value is reduced
        num_color_fields = printing_value
        if len(self.color_entries) > num_color_fields:
            for i in range(num_color_fields, len(self.color_entries)):
                self.color_entries[i].destroy()
                self.color_labels[i].destroy()

            self.color_entries = self.color_entries[:num_color_fields]
            self.color_labels = self.color_labels[:num_color_fields]

        # Configure the submit button and grid layout
        self.submit_button.grid(row=num_color_fields + 13, column=0, columnspan=2, pady=10, sticky="nsew")
        self.main_frame.grid_rowconfigure(num_color_fields + 14, weight=1)
        total_rows = num_color_fields + 15
        self.main_frame.grid_rowconfigure(total_rows, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.submit_button.grid(row=total_rows, column=0, columnspan=2, pady=10, sticky="nsew")

        # Update the window to reflect the changes
        self.window.update()

    def validate_fields(self):
        """
        Validate the input fields for job details.

        Returns:
            bool: True if all fields are valid, False otherwise.

        Comments:
        - Validates the job name, customer name, email address, and mobile number.
        - Validates the width, bottom, height, GSM, and quantity values.
        - Displays appropriate error messages for invalid input.
        - Validates the color details based on the selected printing option.

        """
        
        job_name = self.job_name_entry.get().strip()
        if len(job_name) > 75:
            messagebox.showerror("Error", "Invalid job name. Please enter a valid name (up to 75 characters).")
            return False

        customer_name = self.customer_name_entry.get().strip()
        if not customer_name or len(customer_name) > 75:
            messagebox.showerror("Error", "Invalid customer name. Please enter a valid name (up to 75 characters).")
            return False

        customer_email = self.customer_email_entry.get().strip()
        if not customer_email or not re.match(r"[^@]+@[^@]+\.[^@]+", customer_email):
            messagebox.showerror("Error", "Invalid email address. Please enter a valid email address.")
            return False

        customer_mobile = self.customer_mobile_entry.get().strip()
        if not customer_mobile or not re.match(r"^(?:\+91|0)?[789]\d{9}$", customer_mobile):
            messagebox.showerror("Error", "Invalid mobile number. Please enter a valid 10-digit mobile number starting with 7, 8, or 9.")
            return False

        try:
            width = float(self.width_entry.get())
            if not (5.25 <= width <= 13.00):
                raise ValueError("width")

            bottom = float(self.bottom_entry.get())
            if not (2.5 <= bottom <= 7.00):
                raise ValueError("bottom")

            height = float(self.height_entry.get())
            if not (6.75 <= height <= 17.75):
                raise ValueError("height")

            gsm = float(self.gsm_entry.get())
            if not (55 <= gsm <= 150):
                raise ValueError("gsm")

            quantity = int(self.quantity_entry.get())
            if quantity < 10000:
                raise ValueError("quantity")

        except ValueError as e:
            error_message = "Invalid job details. "
            if str(e) == "width":
                error_message += "Please enter a valid width value within the range of 5.25 to 13.00."
            elif str(e) == "bottom":
                error_message += "Please enter a valid bottom value within the range of 2.5 to 7.00."
            elif str(e) == "height":
                error_message += "Please enter a valid height value within the range of 6.75 to 17.75."
            elif str(e) == "gsm":
                error_message += "Please enter a valid GSM value within the range of 55 to 150."
            else:
                error_message += "Please enter valid quantity within the range of >= 10000."

            messagebox.showerror("Error", error_message)
            return False

        num_colors = int(self.printing_var.get())
        colors = [entry.get().strip() for entry in self.color_entries[:num_colors]]  # Only consider the relevant number of color entries
        if len(colors) != num_colors or any(not color for color in colors):
            messagebox.showerror("Error", "Invalid color details. Please enter a color for each selected printing option.")
            return False

        return True

    def process_data(self):
        """
        Process the entered data and generate a bill of material.

        This method is called when the user submits the form and all fields pass validation.
        It calculates various values based on the entered data, updates an Excel template,
        and saves the modified template as a new bill of material.

        Returns:
            None
        """
        
        if not self.validate_fields():
            return

        # Retrieve data from input fields
        h_in = float(self.height_entry.get())
        b_in = float(self.bottom_entry.get())
        w_in = float(self.width_entry.get())
        gsm = float(self.gsm_entry.get())
        quantity = int(self.quantity_entry.get())
        printing = self.printing_var.get()
        colors = [entry.get().strip() for entry in self.color_entries]

        # Calculate actual height in mm
        act_height = h_in + (b_in / 2) + 1
        act_height_mm = act_height * 25.4

        # Calculate actual width in mm
        act_width_in = 2 * (w_in + b_in) + 1
        act_width_mm = act_width_in * 25.4

        # Choose the closest available cylinder size
        chosen_cylinder = min(AVAILABLE_CYLINDERS, key=lambda x: abs(x - act_height_mm))

        # Choose the immediately smaller and bigger available paper roll sizes
        immediately_smaller_roll = max(
            [x for x in AVAILABLE_PAPER_ROLLS if x <= act_width_mm]
        )
        immediately_bigger_roll = min(
            [x for x in AVAILABLE_PAPER_ROLLS if x > act_width_mm]
        )
        if (act_width_mm - immediately_smaller_roll) <= 5:
            chosen_paper_roll = immediately_smaller_roll
        else:
            chosen_paper_roll = immediately_bigger_roll

        # Calculate various values
        mm_to_m = 0.001
        wpb = (chosen_cylinder * chosen_paper_roll) * (mm_to_m**2) * gsm
        total_finish_weight = wpb * quantity/1000
        total_weight = total_finish_weight / (1 - 0.06)
        side_glue = 0.01 * wpb * quantity / 1000
        bottom_glue = 0.025 * wpb * quantity / 1000
        ink = 0.01 * wpb * quantity / 1000

        # Get the current date and time
        current_date = datetime.datetime.now().strftime("%d/%m/%Y")
        current_time = datetime.datetime.now().strftime("%H:%M:%S")

        # Load the Excel template
        template_path = resource_path(temp_path_env)  
        wb = openpyxl.load_workbook(template_path)
        sheet = wb["Sheet1"]

        # Update the template with the entered data
        sheet["B8"].value = self.generate_job_number()
        sheet["B11"].value = self.customer_name_entry.get()
        sheet["H11"].value = self.customer_email_entry.get()
        sheet["B12"].value = self.customer_mobile_entry.get()
        sheet["B14"].value = self.job_type_combo.get()
        sheet["D15"].value = round(w_in, 2)
        sheet["F15"].value = round(b_in, 2)
        sheet["H15"].value = round(h_in, 2)
        sheet["J15"].value = gsm
        sheet["L15"].value = quantity
        sheet["F8"].value = current_date
        sheet["J8"].value = current_time
        if printing == 0:
            sheet["B16"].value = "No"
        else:
            sheet["B16"].value = "Yes"
        for i, color in enumerate(colors):
            sheet.cell(row=16, column=4 + i * 2).value = color

        # Update the calculated values
        sheet["E19"].value = round(act_height_mm, 2)
        sheet["E20"].value = chosen_cylinder
        sheet["E21"].value = chosen_paper_roll
        sheet["E22"].value = round(wpb, 2)
        sheet["E23"].value = round(total_finish_weight, 2)
        sheet["E24"].value = round(total_weight, 2)
        sheet["E25"].value = round(side_glue, 2)
        sheet["E26"].value = round(bottom_glue, 2)
        sheet["E27"].value = round(ink, 2)

        # Generate a new filename for the bill of material
        new_filename = f"{self.generate_job_number()}.xlsx"
        save_directory = save_directory_env
        backup_directory = backup_directory_env

        # Create the directories if they don't exist
        if not os.path.exists(save_directory):
            os.makedirs(save_directory)

        if not os.path.exists(backup_directory):
            os.makedirs(backup_directory)

        # Define the save and backup paths
        save_path = os.path.join(save_directory, new_filename)
        backup_path = os.path.join(backup_directory, new_filename)

        # Save the modified template as a new bill of material
        wb.save(save_path)
        shutil.copyfile(save_path, backup_path)
        self.handle_fields()
        # Display a success message
        messagebox.showinfo("Success", f"Data saved to {save_path}")

    def generate_job_number(self):
        """
        Generate a unique job number based on the fiscal year and sequential number.

        Returns:
            str: The generated job number.

        Comments:
        - Determines the current date and year.
        - Calculates the fiscal year based on the Indian fiscal cycle.
        - Retrieves the list of files in a directory.
        - Determines the latest job number file based on creation time.
        - Extracts the year and number from the file name.
        - Increments the sequential number if the file year matches the financial year; otherwise, sets it to 1.
        - Constructs the job number in the format: {financial_year}_{sequential_number}.
        """

        current_date = datetime.date.today()
        current_year = current_date.year

        if current_date.month >= 4:  # April or later
            financial_year = f"{str(current_year)[-2:]}-{str(current_year + 1)[-2:]}"
        else:
            financial_year = f"{(current_year - 1)[-2:]}-{str(current_year)[-2:]}"

        files = glob.glob("C:/Backup/*.xlsx")
        latest_job_number_file = max(files, key=os.path.getctime) if files else None

        if latest_job_number_file:
            file_name = os.path.splitext(os.path.basename(latest_job_number_file))[0]
            file_year, file_number = file_name.split("_")

            if file_year == financial_year:
                sequential_number = str(int(file_number) + 1).zfill(7)
            else:
                sequential_number = "0000001"
        else:
            sequential_number = "0000001"

        job_number = f"{financial_year}_{sequential_number}"
        return job_number


    def handle_fields(self):
        """
        Reset the input fields to their default state and update the job number in the GUI.

        Comments:
        - Clears the input fields for job details.
        - Resets the job type combo box selection and printing options.
        - Clears the color fields.
        - Updates the job number in the GUI.
        """

        # Reset fields to default state
        self.job_name_entry.delete(0, "end")
        self.customer_name_entry.delete(0, "end")
        self.customer_email_entry.delete(0, "end")
        self.customer_mobile_entry.delete(0, "end")
        self.job_type_combo.set("Select Job Type")
        self.width_entry.delete(0, "end")
        self.bottom_entry.delete(0, "end")
        self.height_entry.delete(0, "end")
        self.gsm_entry.delete(0, "end")
        self.quantity_entry.delete(0, "end")
        self.printing_var.set(0)
        for entry in self.color_entries:
            entry.delete(0, "end")

        # Clean up the GUI (color fields and labels)
        for label in self.color_labels:
            label.destroy()
        self.color_labels.clear()
        for entry in self.color_entries:
            entry.destroy()
        self.color_entries.clear()

        # Update Job Number in GUI
        self.job_number_label.config(text=f"Job Number: {self.generate_job_number()}")


    def run(self):
        """
        Run the application main loop.

        Comments:
        - Initiates the main event loop for the application window.
        """
        self.window.mainloop()



if __name__ == "__main__":
    estimator = RawMaterialEstimatorGUI()
    estimator.run()