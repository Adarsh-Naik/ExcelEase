from datetime import datetime
import pandas as pd
from fpdf import FPDF
import os
from tkinter import messagebox, Toplevel, Text, Scrollbar, RIGHT, LEFT, Y, BOTH, filedialog, Button
from calendar import monthrange

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



def generate_attendance_pdf(roll_list_file_path, attendance_file_path, subject_excel_path, branch, semester, session, start_date_str, end_date_str, year):
    df = pd.read_excel(attendance_file_path)
    conflicts = check_conflicts(df)

    if conflicts:
        show_conflicts(conflicts)
    else:
        roll_list = pd.read_excel(roll_list_file_path)
        main_data = pd.read_excel(attendance_file_path)
        subject_data = pd.read_excel(subject_excel_path)
    
        subject_names_dict = dict(zip(subject_data['SUBJECT CODE'], subject_data['SUB.']))
    
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d')
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d')
    
        main_data = main_data[(main_data['Date'] >= start_date) & (main_data['Date'] <= end_date)]
        # Your existing logic for generating the attendance report
            # Initialize dictionary to store attendance
        attendance_dict = {}
        
        # Initialize dictionary to store total lectures per subject
        total_lectures_per_subject = {}
        
        # Iterate over each subject
        for subject_code in main_data['Subject Code'].unique():
            # Filter main data for the specific subject and year class
            subject_data = main_data[(main_data['Subject Code'] == subject_code) & (main_data['Class'] == year) & (main_data['Type'] == 'TH') & (main_data['Batch (FOR PR OR Tutorial)'] == 'ALL')]
            
            # Calculate total lectures for the subject
            total_lectures = len(subject_data)
            
            # Store total lectures for the subject
            total_lectures_per_subject[subject_code] = total_lectures
        
        # Define lighter shades of colors for each subject
        lighter_colors = [
            (144, 238, 144),  # Green
            (211, 211, 211),  # Grey
            (255, 200, 100),  # #DE5935
            (173, 216, 230)   # Light Blue
        ]
        
        # Map subjects to their lighter shades
        subject_colors = {}
        for i, subject_code in enumerate(total_lectures_per_subject.keys()):
            if i < len(lighter_colors):
                subject_colors[subject_code] = lighter_colors[i]
            else:
                # Handle case where there are more subjects than colors available
                # Here, you can assign a default color or choose to repeat colors from the beginning of the list
                # For simplicity, let's repeat colors from the beginning of the list
                subject_colors[subject_code] = lighter_colors[i % len(lighter_colors)] 
        
        # Iterate over each student in the roll list
        for index, student_row in roll_list.iterrows():
            # Initialize dictionary to store attendance for each subject
            student_attendance = {}
            total_lectures_attended_all_subjects = 0
            total_lectures_all_subjects = 0
            
            # Filter subjects for the specific year and class of the student
            year_subjects = main_data[(main_data['Class'] == year) & (main_data['Type'] == 'TH') & (main_data['Batch (FOR PR OR Tutorial)'] == 'ALL')]
            
            # Iterate over each subject
            for subject_code in year_subjects['Subject Code'].unique():
                # Filter main data for the specific subject and year class
                subject_data = year_subjects[year_subjects['Subject Code'] == subject_code]    
                
                # Initialize variables to count attendance
                total_lectures_attended = 0
                total_lectures = len(subject_data)
                
                # Iterate over each lecture for the subject
                for _, lecture_row in subject_data.iterrows():
                    # Check if the student's Roll Number is absent for this lecture
                    absent_roll_numbers = str(lecture_row['Absent Numbers']).split(',')  # Assuming 'Absent' is the column containing absent Roll Numbers
                    absent_roll_numbers = [num.strip() for num in absent_roll_numbers]  # Remove leading and trailing spaces
                    
                    if str(student_row['Roll Number']) not in absent_roll_numbers:
                        total_lectures_attended += 1
                
                # Calculate subject attendance percentage
                subject_attendance_percentage = (total_lectures_attended / total_lectures) * 100 if total_lectures != 0 else 0
                
                # Store attendance for the subject
                student_attendance[subject_code] = {'Total Lectures Attended': total_lectures_attended, 'Total Lectures': total_lectures, 'Attendance Percentage': subject_attendance_percentage}
                
                # Update total lectures attended and total lectures for all subjects
                total_lectures_attended_all_subjects += total_lectures_attended
                total_lectures_all_subjects += total_lectures
            
            # Calculate overall attendance percentage for all subjects
            overall_attendance_percentage = (total_lectures_attended_all_subjects / total_lectures_all_subjects) * 100 if total_lectures_all_subjects != 0 else 0
            
            # Store overall attendance for the student
            student_attendance['Overall Attendance'] = overall_attendance_percentage
            
            # Store attendance for the student
            attendance_dict[student_row['Roll Number']] = student_attendance
        
        # Function to add text to PDF
        def add_text(pdf, text, font_size, y_position, align="C"):
            pdf.set_y(y_position)  # Set the Y position
            pdf.set_x(25)  # Adjust the X position as needed
            pdf.set_font("Times", size=font_size)  # Set custom font size
            pdf.cell(0, 10, text, ln=True, align=align)  # Insert the text with the specified alignment
        
        
        # Function to generate PDF
        def generate_pdf1(roll_list, attendance_dict, total_lectures_per_subject, subject_colors, branch_data, main_data, year):
            pdf = FPDF(orientation='P')
            pdf.add_page()
            
            # Add image at the top left of the first page
            pdf.image("college_logo.png", x=10, y=10, w=30)  # Adjust the path and dimensions accordingly
        
        
            add_text(pdf, "GOVERNMENT COLLEGE OF ENGINEERING, YAVATMAL", 10, 5)
            add_text(pdf, "Affiliated to Dr. Babasaheb Ambedkar Technological University, Lonere", 8, 10)
            add_text(pdf, f"Branch: {branch_data}", 10, 15)
            add_text(pdf, f"Even Sem, Session {session}", 10, 20)
            add_text(pdf, f"{start_date.strftime('%B')} Month Attendance, {year}, {semester} semester", 10, 25)
            add_text(pdf, f"Attendance From: {start_date.strftime('%d-%m-%Y')} To: {end_date.strftime('%d-%m-%Y')}", 10, 30)
            pdf.set_font("Times", size=7)
            pdf.ln(15)  # Move to the next line
            # Add subject code headers above the table
            pdf.set_fill_color(244, 166, 255)
            pdf.set_text_color(0, 0, 0)
            pdf.cell(25, 5, "", 0, 0)  # Add an empty cell to align with the Roll No. column
            pdf.cell(30, 5, "Subject Codes -->", 0, 0)
            for subject_code in attendance_dict[next(iter(attendance_dict))].keys():
               if subject_code != 'Overall Attendance':
                  subject_name = subject_names_dict.get(subject_code, "Unknown")
                  subject_header = f"{subject_name}({subject_code})"
                  pdf.cell(20, 5, subject_header, 1, 0, 'C', 1)
            
            # Add a header
            pdf.ln()  # Move to the next line
            pdf.cell(10, 5, "Roll No.", 1, 0, 'C', 1)
            pdf.cell(45, 5, "Name of Student", 1, 0, 'C', 1)
            
            subject_codes = [code for code in attendance_dict[next(iter(attendance_dict))].keys() if code != 'Overall Attendance']
            
            # Add individual attendance percentages headers
            for subject_code in attendance_dict[next(iter(attendance_dict))].keys():
                if subject_code != 'Overall Attendance':
                    pdf.cell(10, 5, f"Out of {total_lectures_per_subject[subject_code]}", 1, 0, 'C', 1)
                    pdf.cell(10, 5, "%", 1, 0, 'C', 1)
            
            pdf.cell(20, 5, "Overall Attendance", 1, 1, 'C', 1)
            
            # Add data rows
            for roll_number, attendance_data in attendance_dict.items():
                name_of_student = roll_list.loc[roll_list['Roll Number'] == roll_number, 'Name of Student'].iloc[0]
                pdf.cell(10, 5, str(roll_number), 1, 0, 'C')
                pdf.cell(45, 5, str(name_of_student), 1, 0, 'L')
                
                for subject_code, data in attendance_data.items():
                    if subject_code != 'Overall Attendance':
                        # Determine the color for the subject
                        color = subject_colors.get(subject_code, (255, 255, 255))  # Default to white if subject code not found in subject_colors
                        pdf.set_fill_color(*color)
                        if data['Attendance Percentage'] < 75:
                            # If attendance is less than 75%, use red color
                            pdf.set_fill_color(255, 102, 102)  # Red color
                        pdf.cell(10, 5, f"{data['Total Lectures Attended']}", 1, 0, 'C', fill=True)
                        pdf.cell(10, 5, "{:.2f}%".format(data['Attendance Percentage']), 1, 0, 'C', fill=True)
                
                # Determine overall attendance color
                overall_attendance_percentage = attendance_data['Overall Attendance']
                overall_color = (255, 255, 255)  # Default white color
                if overall_attendance_percentage < 75:
                    overall_color = (255, 102, 102)  # Red color for less than 75%
                pdf.set_fill_color(*overall_color)
                pdf.cell(20, 5, "{:.2f}%".format(overall_attendance_percentage), 1, 1, 'C', fill=True)
                
            pdf.cell(55, 10, "Sign of Subject Incharge", 1, 0, 'C', 1)   
            for subject_code in subject_codes:
                pdf.cell(20, 10, " ", 1, 0, 'C')
                
            pdf.ln()  
            
            pdf.cell(55, 5, " Name of Subject Incharge", 1, 0, 'C', 1)   
            for subject_code in subject_codes:
                faculty_name = main_data.loc[main_data['Subject Code'] == subject_code, 'Name of Faculty'].iloc[0]
                pdf.cell(20, 5, str(faculty_name), 1, 0, 'C')
        
            pdf.ln(40)
           
        
            # Row for names under the signatures
            pdf.set_font("Times", style='B', size=7)  # Set font to bold and size 7
            pdf.cell(75, 5, "Class Incharge", 0, 0, 'C')  # Center aligned, no border
            year_name = f" {year}"  # Example name, replace with actual variable or lookup if needed
            pdf.cell(150, 5, "Head Of Department", 0, 1, 'C')  # Center aligned, no border
        
            # Additional cell for name of HOD
            branch = f"{branch_data} Department"  # Example name, replace with actual variable or lookup if needed
            pdf.cell(75, 5, year_name, 0, 0, 'C')  # Center aligned, no border
            pdf.cell(150, 5, branch, 0, 1, 'C')  # Center aligned, no border
        
            # Save the PDF to a file
            # pdf.output("Attendance Report New11.pdf")
            # os.startfile("Attendance Report New11.pdf")
    
            home_directory = os.path.expanduser('~')
            output_folder = os.path.join(home_directory, 'ExcelEase', 'Attendance')
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)
    
            # Define output file path
            output_file = os.path.join(output_folder, f"{session}_{year}_{semester}.pdf")
            pdf.output(output_file)
    
            messagebox.showinfo("Success", f"PDF generated successfully: {output_file}")
    
        generate_pdf1(roll_list, attendance_dict, total_lectures_per_subject, subject_colors, branch ,main_data, year)
    