from datetime import datetime
import pandas as pd
from fpdf import FPDF
import os
from tkinter import messagebox, Toplevel, Text, Scrollbar, RIGHT, LEFT, Y, BOTH, filedialog, Button
from calendar import monthrange
from fpdf import FPDF

import pandas as pd
    

def check_conflicts(df):
    df['Date'] = pd.to_datetime(df['Date'], format='%m-%d-%Y')
    df['Start Time'] = pd.to_datetime(df['Start Time'], format='%H:%M:%S').dt.time

    conflicts = []

    # Check for faculty conflicts
    faculties = df['Name of Faculty'].unique()
    for faculty in faculties:
        faculty_data = df[df['Name of Faculty'] == faculty]
        grouped = faculty_data.groupby(['Date', 'Start Time'])

        for name, group in grouped:
            if len(group) > 1:
                conflicts.append({
                    'Type': 'Faculty Conflict',
                    'Faculty': faculty,
                    'Date': name[0].strftime('%m-%d-%Y'),
                    'Time': name[1].strftime('%H:%M:%S'),
                    'Conflicts': group[['Class', 'Subject Code', 'Type', 'Batch (FOR PR OR Tutorial)']].values.tolist()
                })

    # Check for class conflicts
    classes = df['Class'].unique()
    for class_name in classes:
        class_data = df[df['Class'] == class_name]
        grouped = class_data.groupby(['Date', 'Start Time', 'Batch (FOR PR OR Tutorial)'])

        for name, group in grouped:
            if len(group) > 1:
                conflicts.append({
                    'Type': 'Class Conflict',
                    'Class': class_name,
                    'Date': name[0].strftime('%m-%d-%Y'),
                    'Time': name[1].strftime('%H:%M:%S'),
                    'Batch': name[2],
                    'Conflicts': group[['Name of Faculty', 'Subject Code', 'Type']].values.tolist()
                })

    return conflicts

def save_conflicts_to_pdf(conflicts, filename):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.set_font("Arial", size=10)

    col_width = pdf.w / 4.5  # Column width
    row_height = pdf.font_size * 1.5  # Row height

    for index, conflict in enumerate(conflicts, start=1):
        pdf.set_font("Arial", style='B', size=10)
        pdf.cell(col_width, row_height, f"Conflict {index}:", border=1)
        pdf.set_font("Arial", size=10)
        pdf.cell(col_width, row_height, f"Type: {conflict['Type']}", border=1)
        if conflict['Type'] == 'Faculty Conflict':
            pdf.cell(col_width, row_height, f"Faculty: {conflict['Faculty']}", border=1,ln=True)
        else:
            pdf.cell(col_width, row_height, f"Class: {conflict['Class']}", border=1)
            pdf.cell(col_width, row_height, f"Batch: {conflict['Batch']}", border=1, ln=True)
        pdf.cell(col_width, row_height, f"Date: {conflict['Date']}", border=1)
        pdf.cell(col_width, row_height, f"Time: {conflict['Time']}", border=1, ln=True)
        
        pdf.cell(col_width, row_height, "Conflicts:", border=1, ln=True)
        for detail in conflict['Conflicts']:
            pdf.cell(0, row_height, f"  - {', '.join(map(str, detail))}", border=1, ln=True)

        pdf.cell(0, row_height, "", ln=True)  # Blank line for spacing

    pdf.output(filename)

def show_conflicts(conflicts):
    conflict_window = Toplevel()
    conflict_window.title("Conflicts Found")
    
    text_widget = Text(conflict_window)
    scrollbar = Scrollbar(conflict_window, command=text_widget.yview)
    
    text_widget.configure(yscrollcommand=scrollbar.set)
    
    text_widget.pack(side=LEFT, fill=BOTH, expand=True)
    scrollbar.pack(side=RIGHT, fill=Y)
    
    # Function to save conflicts to PDF
    def save_as_pdf():
        filename = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if filename:
            save_conflicts_to_pdf(conflicts, filename)

    Button(conflict_window, text="Save as PDF", command=save_as_pdf).pack()
    
    for index, conflict in enumerate(conflicts, start=1):
        message = f"Conflict {index}:\n"
        message += f"Conflict Type: {conflict['Type']}\n"
        if conflict['Type'] == 'Faculty Conflict':
            message += f"Faculty: {conflict['Faculty']}\n"
        else:
            message += f"Class: {conflict['Class']}\n"
            message += f"Batch: {conflict['Batch']}\n"
        message += f"Date: {conflict['Date']}\n"
        message += f"Time: {conflict['Time']}\n"
        message += "Conflicts:\n"
        for detail in conflict['Conflicts']:
            message += f"  - {detail}\n"
        message += "\n"
        
        text_widget.insert('end', message)
    
    text_widget.configure(state='disabled')




def generate_faculty_bill(file_path, branch, subject_file_path, faculty_name, first_year_roll_list_path, second_year_roll_list_path, third_year_roll_list_path, final_year_roll_list_path, selected_month, theory_amount, other_amount):
    try:
        df = pd.read_excel(file_path)
        conflicts = check_conflicts(df)

        if conflicts:
            show_conflicts(conflicts)
        else:
            # Generate PDF
            selected_month_date = datetime.strptime(selected_month, "%B %Y")
            selected_month_num = selected_month_date.month
            selected_year = selected_month_date.year
            print("function started", theory_amount, other_amount)
            theory_amount = int(theory_amount)
            other_amount = int(other_amount)
    
            first_year_roll_list = pd.read_excel(first_year_roll_list_path)
            second_year_roll_list = pd.read_excel(second_year_roll_list_path)
            third_year_roll_list = pd.read_excel(third_year_roll_list_path)
            final_year_roll_list = pd.read_excel(final_year_roll_list_path)
    
            # Verify that the roll lists have the correct lengths
            for roll_list, class_name in zip(
                [first_year_roll_list, second_year_roll_list, third_year_roll_list, final_year_roll_list],
                ["First Year", "Second Year", "Third Year", "Final Year"]
            ):
                max_roll_number = roll_list['Roll Number'].max()
                if len(roll_list) != max_roll_number:
                    raise ValueError(f"Roll list for {class_name} has inconsistent length. Expected {max_roll_number}, but got {len(roll_list)}.")
    
            # Read the Excel file
            df = pd.read_excel(file_path)
    
            # Read the subject Excel file
            subject_df = pd.read_excel(subject_file_path)
    
            # Merge subject information with the main data frame
            df = pd.merge(df, subject_df, left_on='Subject Code', right_on='SUBJECT CODE', how='left')
    
            # Convert 'Date' column to datetime
            df['Date'] = pd.to_datetime(df['Date'], format='%d-%m-%Y')
    
            # Extract month and year from 'Date' column
            df['Month'] = df['Date'].dt.month
            df['Year'] = df['Date'].dt.year
    
            # Filter data based on faculty name and selected month
            faculty_data = df[(df['Name of Faculty'] == faculty_name) & (df['Month'] == selected_month_num) & (df['Year'] == selected_year)]
    
            # Group filtered data by Class, subject code, lecture type, year, and month
            grouped_data = faculty_data.groupby(['Class', 'Subject Code', 'Type', 'Year', 'Month'])
    
            # Sort grouped data by subject code and date in ascending order
            grouped_data = sorted(grouped_data, key=lambda x: (x[0][1], x[0][4], x[0][3]))
    
            # Generate PDF
            pdf = FPDF(orientation='L', unit='mm', format='legal')
            total_bill = 0  # Initialize total bill
            
            # Dictionary to store subject-wise total amounts and lecture counts
            subject_totals = {}
            all_subjects = []
    
            for (Class_type, subject_code, lecture_type, year, Month_no), data in grouped_data:
                # Select appropriate roll list based on class year
                if Class_type == 'First Year':
                    roll_list = first_year_roll_list
                elif Class_type == 'Second Year':
                    roll_list = second_year_roll_list
                elif Class_type == 'Third Year':
                    roll_list = third_year_roll_list
                elif Class_type == 'Final Year':
                    roll_list = final_year_roll_list
                else:
                    continue
    
                # Start a new page for each subject and lecture type
                pdf.add_page()
    
                # Get the subject name
                subject_name = data.iloc[0]['SUBJECT']
    
                # Set font for title
                pdf.set_font("Times", size=12)
    
                # Add Name of Faculty, Year, and Month as title
                Month_no_str = datetime.strptime(str(Month_no), "%m").strftime("%B")
                start_date = datetime(year, Month_no, 1).strftime('%d-%m-%Y')
                end_date = datetime(year, Month_no, monthrange(year, Month_no)[1]).strftime('%d-%m-%Y')
                pdf.cell(340, 10, f"Name of Faculty : {faculty_name}", ln=True, align='C')
                pdf.cell(340, 10, f"Month : {Month_no_str} {year}", ln=True, align='C')
                pdf.cell(340, 10, f"Year : {Class_type}", ln=True, align='C')
                pdf.cell(340, 10, f"Subject Code : {subject_code} - {subject_name} ({lecture_type})", ln=True, align='C')
    
                pdf.ln(7)  # Add empty line
    
                # Set font for table content
                pdf.set_font("Times", size=10)
    
                # Add table headers
                pdf.cell(10, 10, 'Sr.No.', 1, 0, 'C')
                pdf.cell(20, 10, 'Date', 1, 0, 'C')
                pdf.cell(20, 10, 'Day', 1, 0, 'C')
                pdf.cell(20, 10, 'Time', 1, 0, 'C')
                pdf.cell(20, 10, 'Branch', 1, 0, 'C')
                pdf.cell(20, 10, 'Subject', 1, 0, 'C')
                pdf.cell(20, 10, 'TH/PR', 1, 0, 'C')
                pdf.cell(20, 10, 'Batch', 1, 0, 'C')
                pdf.cell(30, 10, 'No. of Students', 1, 0, 'C')
                pdf.cell(85, 10, 'Topic Name', 1, 0, 'C')
                pdf.cell(25, 10, 'Sign of Faculty', 1, 0, 'C')
                pdf.cell(25, 10, 'Sign of Verifier', 1, 0, 'C')
                pdf.cell(25, 10, 'Sign of HOD', 1, 0, 'C')
                pdf.ln()  # Move to the next line
    
                # Initialize serial number
                serial_number = 1
    
                # Initialize bill for the lecture type
                bill = 0
    
                # Initialize lecture count for the lecture type
                lecture_count = 0
                
                # Sort data by date in ascending order
                data = data.sort_values(by='Date')
    
                # Add table rows
                for index, row in data.iterrows():
                    try:
                        # Convert string date to datetime object
                        date_str = pd.to_datetime(row['Date']).strftime('%d-%m-%y')
                        day_of_week = pd.to_datetime(row['Date']).strftime('%A')
    
                        # Calculate number of students present
                        batch = row['Batch (FOR PR OR Tutorial)']
                        if batch == 'ALL':
                            total_students = len(roll_list)
                        else:
                            # Split the batch string if it contains multiple entries
                            batches = [b.strip() for b in batch.split(',')]
                            # Initialize total_students for this batch
                                                # Initialize total_students for this batch
                            total_students = 0
                            for b in batches:
                               # Calculate the total number of students in each batch and add it to total_students
                               total_students += len(roll_list[roll_list['Batch'] == b])
    
                        absent_students = []
                        if pd.notnull(row['Absent Numbers']):
                           absent_students = [num.strip() for num in str(row['Absent Numbers']).split(',') if num.strip()]
     
                        num_absent_students = len(absent_students) if row['Absent Numbers'] else 0
                        num_students_present = total_students - num_absent_students
    
                        pdf.cell(10, 10, str(serial_number), 1, 0, 'C')  # Add serial number
                        pdf.cell(20, 10, date_str, 1, 0, 'C')
                        pdf.cell(20, 10, day_of_week, 1, 0, 'C')  # Use the variable day_of_week instead of row['Date'].strftime('%A')
                        pdf.cell(20, 10, str(row['Start Time']), 1, 0, 'C')
                        pdf.cell(20, 10, branch, 1, 0, 'C')
                        pdf.cell(20, 10, str(row['SUB.']), 1, 0, 'C')  # Add subject name
                        pdf.cell(20, 10, str(row['Type']), 1, 0, 'C')
                        pdf.cell(20, 10, str(batch), 1, 0, 'C')
                        pdf.cell(30, 10, str(num_students_present), 1, 0, 'C')
                        pdf.cell(85, 10, '', 1, 0, 'C')
                        pdf.cell(25, 10, '', 1, 0, 'C')
                        pdf.cell(25, 10, '', 1, 0, 'C')
                        pdf.cell(25, 10, '', 1, 0, 'C')
                       
                        pdf.ln()  # Move to the next line
                        serial_number += 1  # Increment serial number
    
                        # Calculate bill for the lecture
                        if lecture_type == 'TH':
                            bill_per_lecture = theory_amount
                            lecture_count += 1
                        else:
                            bill_per_lecture = other_amount
                            lecture_count += 1
    
                        bill += bill_per_lecture  # Update total bill
    
                    except Exception as e:
                        print(f"Error processing row {index}: {e}")
    
                # Add bill to total bill
                total_bill += bill
                
                # Update subject-wise total amounts and lecture counts
                subject_total_key = f"{subject_code} {lecture_type}"
                if subject_total_key not in subject_totals:
                    subject_totals[subject_total_key] = {'total_amount': 0, 'lecture_count': 0}
                subject_totals[subject_total_key]['total_amount'] += bill
                subject_totals[subject_total_key]['lecture_count'] += lecture_count
                subject_name = data.iloc[0]['SUBJECT']
    
                all_subjects.append(f"{subject_name} ({subject_code}) - {lecture_type}")
    
                pdf.ln(5)
    
                # Add short summary of the particular subject
                pdf.set_font("Times", size=12)
                pdf.cell(0, 10, f"Subject Summary:- {subject_code} {lecture_type} = Total Slots - {lecture_count} Hrs = Total Amount - Rs.{bill}", ln=True, align='L')
    
            # Add summary page
            pdf.add_page()
    
            # Add subject-wise summary at the end
            pdf.ln(10)
            pdf.set_font("Times", size=12)
            # Update the certification text line with dynamic dates and all subjects taught
            certification_text = f'                 It is certified that from {start_date} to {end_date} in Computer Engineering Department {", ".join(all_subjects)} subjects have been taken and its details are as follows :'
            max_line_length = 163  # Define the maximum line length
            # Split the certification text into multiple lines if it exceeds the maximum line length
            certification_lines = [certification_text[i:i+max_line_length] for i in range(0, len(certification_text), max_line_length)]
            for line in certification_lines:
                pdf.cell(0, 10, line, ln=True, align='L')
    
            pdf.ln(5)
            for subject, details in subject_totals.items():
                pdf.cell(0, 10, f"{subject}  :  Total Lectures - {details['lecture_count']}, Total Amount = Rs.{details['total_amount']} ", ln=True, align='L')
            
            # Calculate and add total of all subject-wise amounts
            total_all_subjects_amount = sum(details['total_amount'] for details in subject_totals.values())
            pdf.cell(0, 10, f"Total of All Subject-wise Amounts: Rs. {total_all_subjects_amount}", ln=True, align='L')
    
            # Calculate and add grand total of all subject-wise amounts
            grand_total = sum(details['total_amount'] for details in subject_totals.values())
            pdf.cell(0, 10, f"Grand Total of All Subject-wise Amounts: Rs. {grand_total}", ln=True, align='L')
    
    
                # Define output folder in user's home directory and create it if it doesn't exist
            home_directory = os.path.expanduser('~')
            output_folder = os.path.join(home_directory, 'ExcelEase', 'Bills')
            if not os.path.exists(output_folder):
               os.makedirs(output_folder)
            # Define output file path
            output_file = os.path.join(output_folder, f"{faculty_name}_bill_{selected_month}.pdf")
            pdf.output(output_file)
            messagebox.showinfo("Success", f"PDF generated successfully: {output_file}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


# Now you can call the `generate_faculty_bill` function separately from other functions.

