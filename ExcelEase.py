import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog
from tkcalendar import DateEntry
import pandas as pd
from PIL import Image, ImageTk, ImageFilter
import os
import time
import configparser
import shutil
import generate_bill 
import generate_attendance
from tkinter import messagebox
# BLUE

def typing_animation(label, text, delay=0.02):
    """Simulate typing animation for a label."""
    for char in text:
        label.config(text=label.cget("text") + char)
        label.update()
        time.sleep(delay)
        

class StartWindow(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Excel Ease")
        self.geometry("1200x700")
        self.config(bg="#38B6FF")

        #00ABE1  light 
        #161F6D  dark
        
        # Create a label for typing animation
        typing_label = tk.Label(self, font=("Gill Sans Ultra Bold", 36),bg="#38B6FF",fg="#161F6D")
        typing_label.place(x=300,y=130)
        text_to_type = "Welcome."
        typing_animation(typing_label, text_to_type)
        
        typing_label = tk.Label(self, font=("Gill Sans Ultra Bold", 36),bg="#38B6FF",fg="white")
        typing_label.place(x=610,y=130)
        text_to_type = "To"
        typing_animation(typing_label, text_to_type)
        
        typing_label = tk.Label(self, font=("Gill Sans Ultra Bold", 36),bg="#38B6FF",fg="#161F6D")
        typing_label.place(x=700,y=130)
        text_to_type = "The"
        typing_animation(typing_label, text_to_type)

        typing_label = tk.Label(self, font=("Gill Sans Ultra Bold", 36),bg="#38B6FF",fg="white")
        typing_label.place(x=300,y=200)
        text_to_type = "Government"
        typing_animation(typing_label, text_to_type)
        
        typing_label = tk.Label(self, font=("Gill Sans Ultra Bold", 36),bg="#38B6FF",fg="#161F6D")
        typing_label.place(x=300,y=270)
        text_to_type = "College Of"
        typing_animation(typing_label, text_to_type)

        # Load and resize the image
        image_path = "circle.png"  # Replace "your_image.png" with your image file path
        original_image = Image.open(image_path)
        resized_image = original_image.resize((270, 270), Image.ANTIALIAS)  # Customize the size as needed

        # Convert the image for tkinter
        self.image = ImageTk.PhotoImage(resized_image)

        # Create a label to display the image
        self.image_label = tk.Label(self, image=self.image,borderwidth=0)
        self.image_label.place(x=700,y=220)

        typing_label = tk.Label(self, font=("Gill Sans Ultra Bold", 36),bg="#38B6FF",fg="white")
        typing_label.place(x=300,y=340)
        text_to_type = "Engineering"
        typing_animation(typing_label, text_to_type)

        typing_label = tk.Label(self, font=("Gill Sans Ultra Bold", 36),bg="#38B6FF",fg="#161F6D")
        typing_label.place(x=670,y=340)
        text_to_type = ","
        typing_animation(typing_label, text_to_type)

        typing_label = tk.Label(self, font=("Showcard Gothic", 48),bg="#38B6FF",fg="#161F6D")
        typing_label.place(x=300,y=410)
        text_to_type = "Yavatmal"
        typing_animation(typing_label, text_to_type)

        
        self.get_started_button = tk.Button(self, text="Get Started",font=("Segoe UI Black",22), command=self.open_second_window,bg="white",fg="#161F6D",borderwidth=0)
        self.get_started_button.pack(pady=10)
        self.get_started_button.place(x=310,y=520,width=620,height=40)
         
        typing_label = tk.Label(self, font=("Sitka Small Semibold", 10),bg="#38B6FF",fg="#161F6D")
        typing_label.place(x=550,y=670)
        text_to_type = "@ Developed By_________"
        typing_animation(typing_label, text_to_type)

        typing_label = tk.Label(self, font=("Sitka Small Semibold", 10),bg="#38B6FF",fg="white")
        typing_label.place(x=720,y=670)
        text_to_type = "Pankaj Gite, Prajwal Ghugare, Adarsh Naik & Roshan Tajane."
        typing_animation(typing_label, text_to_type) 
    
            
    
    def open_second_window(self):
        self.destroy()
        second_window = SecondWindow()
        second_window.mainloop()


class SecondWindow(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Excel Ease")
        self.config(bg="white")
        self.geometry("1098x730")
        self.resizable(False,False)
        #self.load_config()

        # Load the image
        image = Image.open("bannerMain.png")

        # Resize the image
        width, height = 1100, 175
        image = image.resize((width, height), Image.ANTIALIAS)

        # Convert the Image object into a Tkinter PhotoImage object
        photo = ImageTk.PhotoImage(image)

        # Create a label and display the image
        self.image_label = tk.Label(self, image=photo,borderwidth=0)
        self.image_label.image = photo  # Keep a reference to the image to prevent garbage collection
        self.image_label.pack()

        # Load the image
        image = Image.open("logo.png")
        # Resize the image
        width, height = 150, 150
        image = image.resize((width, height), Image.ANTIALIAS)
        # Convert the Image object into a Tkinter PhotoImage object
        photo = ImageTk.PhotoImage(image)
        # Create a label and display the image
        self.image_label = tk.Label(self, image=photo,borderwidth=0)
        self.image_label.image = photo  # Keep a reference to the image to prevent garbage collection
        self.image_label.pack()
        self.image_label.place(x=50,y=15)

        # Frame to contain the buttons
        button_frame = tk.Frame(self)
        button_frame.pack(side=tk.TOP, fill=tk.X)
        button_frame.config(bg="white")

        self.generate_bill_button = tk.Button(button_frame, text="Generate Bill",font=("Verdana",14), command=self.show_generate_bill_form,bg="#161F6D",fg="white")
        self.generate_bill_button.pack(side=tk.LEFT)
        self.generate_bill_button.config(width=30,height=2,borderwidth=0)

        self.attendance_report_button = tk.Button(button_frame, text="Attendance Report", command=self.show_attendance_report_form,bg="white",fg="#161F6D",font=("Verdana",14))
        self.attendance_report_button.pack(side=tk.LEFT)
        self.attendance_report_button.config(width=30,height=2,borderwidth=0)

        self.change_configuration_button = tk.Button(button_frame, text="Change Configuration", command=self.show_change_configuration_form,bg="white",fg="#161F6D",font=("Verdana",14))
        self.change_configuration_button.pack(side=tk.LEFT)
        self.change_configuration_button.config(width=30,height=2,borderwidth=0)

        # Create a frame for the form with background color and border
        self.form_frame = tk.Frame(self, bg="#DDEDF4", padx=20, pady=20, relief=tk.RAISED, borderwidth=0)
        self.form_frame.place(x=10,y=100,width=600,height=250, relx=0.5, rely=0.5, anchor=tk.CENTER)

        # File selection
        self.file_label = tk.Label(self.form_frame, text="Select Excel File:", bg="#DDEDF4", fg="#161F6D", font=("Franklin Gothic Demi", 12))
        self.file_label.grid(row=0, column=0, pady=10)

        self.file_entry = tk.Entry(self.form_frame)
        self.file_entry.config(bg="white")
        self.file_entry.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        self.file_entry.place(x=140,y=10,height=25,width=290)

        self.file_button = tk.Button(self.form_frame, text="Browse", bg="#161F6D", fg="white", borderwidth=0, width=12, font=("Franklin Gothic Demi", 10), command=self.select_file)
        self.file_button.grid(row=0, column=2, pady=10, padx=15)

        # Faculty selection
        self.faculty_label = tk.Label(self.form_frame, text="Select Faculty:", bg="#DDEDF4", fg="#161F6D", font=("Franklin Gothic Demi", 12))
        self.faculty_label.grid(row=1, column=0, pady=10)

        self.faculty_var = tk.StringVar()
        self.faculty_dropdown = tk.OptionMenu(self.form_frame, self.faculty_var, '')
        self.faculty_dropdown.config(width=40, bg="white", borderwidth=0)
        self.faculty_dropdown.grid(row=1, column=1, padx=10, pady=10, sticky="w")

        # Month selection
        self.month_label = tk.Label(self.form_frame, text="Select Month:", bg="#DDEDF4", fg="#161F6D", font=("Franklin Gothic Demi", 12))
        self.month_label.grid(row=2, column=0, pady=10)

        self.month_var = tk.StringVar()
        self.month_checkboxes_frame = tk.Frame(self.form_frame, bg="#DDEDF4")
        self.month_checkboxes_frame.grid(row=2, column=1, sticky="w", padx=10, pady=10)

        # Branch selection
        self.branch_label = tk.Label(self.form_frame, text="Select Branch:",bg="#DDEDF4",fg="#161F6D",font=("Franklin Gothic Demi",12))
        self.branch_label.grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.branch_combobox = ttk.Combobox(self.form_frame, values=["Computer", "Mechanical", "Electrical", "E&TC", "Civil"],width=34)
        self.branch_combobox.grid(row=3, column=1, padx=10, pady=10)

        #Generate button
        self.generate_bill_button = tk.Button(self, text="Generate Bill",bg="#161F6D",fg="white",font=("Dubai Medium-Bold",14), command=self.call_generate_bill)
        self.generate_bill_button.pack(side=tk.LEFT)
        self.generate_bill_button.config(height=2,borderwidth=0)
        self.generate_bill_button.place(x=258,y=580,width=600)


    def show_generate_bill_form(self):
        self.clear_widgets()
        self.config(bg="white")

        # Load the image
        image = Image.open("bannerMain.png")

        # Resize the image
        width, height = 1100, 175
        image = image.resize((width, height), Image.ANTIALIAS)

        # Convert the Image object into a Tkinter PhotoImage object
        photo = ImageTk.PhotoImage(image)

        # Create a label and display the image
        self.image_label = tk.Label(self, image=photo,borderwidth=0)
        self.image_label.image = photo  # Keep a reference to the image to prevent garbage collection
        self.image_label.pack()


        # Load the image
        image = Image.open("logo.png")
        # Resize the image
        width, height = 150, 150
        image = image.resize((width, height), Image.ANTIALIAS)
        # Convert the Image object into a Tkinter PhotoImage object
        photo = ImageTk.PhotoImage(image)
        # Create a label and display the image
        self.image_label = tk.Label(self, image=photo,borderwidth=0)
        self.image_label.image = photo  # Keep a reference to the image to prevent garbage collection
        self.image_label.pack()
        self.image_label.place(x=50,y=15)

        # Frame to contain the buttons
        button_frame = tk.Frame(self)
        button_frame.pack(side=tk.TOP, fill=tk.X)
        button_frame.config(bg="white")


        self.generate_bill_button = tk.Button(button_frame, text="Generate Bill",font=("Verdana",14), command=self.show_generate_bill_form,bg="#161F6D",fg="white")
        self.generate_bill_button.pack(side=tk.LEFT)
        self.generate_bill_button.config(width=30,height=2,borderwidth=0)

        self.attendance_report_button = tk.Button(button_frame, text="Attendance Report", command=self.show_attendance_report_form,bg="white",fg="#161F6D",font=("Verdana",14))
        self.attendance_report_button.pack(side=tk.LEFT)
        self.attendance_report_button.config(width=30,height=2,borderwidth=0)

        self.change_configuration_button = tk.Button(button_frame, text="Change Configuration", command=self.show_change_configuration_form,bg="white",fg="#161F6D",font=("Verdana",14))
        self.change_configuration_button.pack(side=tk.LEFT)
        self.change_configuration_button.config(width=30,height=2,borderwidth=0)

    # def create_widgets(self):

        # Create a frame for the form with background color and border
        self.form_frame = tk.Frame(self, bg="#DDEDF4", padx=20, pady=20, relief=tk.RAISED, borderwidth=0)
        self.form_frame.place(x=10,y=100,width=600,height=250, relx=0.5, rely=0.5, anchor=tk.CENTER)

        # File selection
        self.file_label = tk.Label(self.form_frame, text="Select Excel File:", bg="#DDEDF4", fg="#161F6D", font=("Franklin Gothic Demi", 12))
        self.file_label.grid(row=0, column=0, pady=10)

        self.file_entry = tk.Entry(self.form_frame)
        self.file_entry.config(bg="white")
        self.file_entry.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        self.file_entry.place(x=140,y=10,height=25,width=290)

        self.file_button = tk.Button(self.form_frame, text="Browse", bg="#161F6D", fg="white", borderwidth=0, width=12, font=("Franklin Gothic Demi", 10), command=self.select_file)
        self.file_button.grid(row=0, column=2, pady=10, padx=15)
        
        
        # Faculty selection
        self.faculty_label = tk.Label(self.form_frame, text="Select Faculty:", bg="#DDEDF4", fg="#161F6D", font=("Franklin Gothic Demi", 12))
        self.faculty_label.grid(row=1, column=0, pady=10)

        self.faculty_var = tk.StringVar()
        self.faculty_dropdown = tk.OptionMenu(self.form_frame, self.faculty_var, '')
        self.faculty_dropdown.config(width=40, bg="white", borderwidth=0)
        self.faculty_dropdown.grid(row=1, column=1, padx=10, pady=10, sticky="w")

        # Month selection
        self.month_label = tk.Label(self.form_frame, text="Select Month:", bg="#DDEDF4", fg="#161F6D", font=("Franklin Gothic Demi", 12))
        self.month_label.grid(row=2, column=0, pady=10)

        self.month_var = tk.StringVar()
        self.month_checkboxes_frame = tk.Frame(self.form_frame, bg="#DDEDF4")
        self.month_checkboxes_frame.grid(row=2, column=1, sticky="w", padx=10, pady=10)

        # Branch selection
        self.branch_label = tk.Label(self.form_frame, text="Select Branch:",bg="#DDEDF4",fg="#161F6D",font=("Franklin Gothic Demi",12))
        self.branch_label.grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.branch_combobox = ttk.Combobox(self.form_frame, values=["Computer", "Mechanical", "Electrical", "E&TC", "Civil"],width=34)
        self.branch_combobox.grid(row=3, column=1, padx=10, pady=10)

        #Generate button
        self.generate_bill_button = tk.Button(self, text="Generate Bill",bg="#161F6D",fg="white",font=("Dubai Medium-Bold",14), command=self.call_generate_bill)
        self.generate_bill_button.pack(side=tk.LEFT)
        self.generate_bill_button.config(height=2,borderwidth=0)
        self.generate_bill_button.place(x=258,y=580,width=600)

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        self.file_entry.delete(0, tk.END)
        self.file_entry.insert(0, file_path)

        if file_path:
            # Read faculty names and available months
            df = pd.read_excel(file_path)
            faculty_names = df['Name of Faculty'].unique().tolist()
            available_months = pd.to_datetime(df['Date']).dt.strftime('%B %Y').unique().tolist()
            self.faculty_var.set('')
            self.faculty_dropdown['menu'].delete(0, 'end')
            for name in faculty_names:
                self.faculty_dropdown['menu'].add_command(label=name, command=lambda n=name: self.faculty_var.set(n))
            for month in available_months:
                month_radio = tk.Radiobutton(self.month_checkboxes_frame, text=month, variable=self.month_var, value=month, bg="#DDEDF4")
                month_radio.grid(sticky="w", padx=10, pady=(0, 5))
    
    def show_attendance_report_form(self):
        self.clear_widgets()
        self.config(bg="white")

        # Load the image
        image = Image.open("bannerMain.png")

        # Resize the image
        width, height = 1100, 175
        image = image.resize((width, height), Image.ANTIALIAS)

        # Convert the Image object into a Tkinter PhotoImage object
        photo = ImageTk.PhotoImage(image)

        # Create a label and display the image
        self.image_label = tk.Label(self, image=photo,borderwidth=0)
        self.image_label.image = photo  # Keep a reference to the image to prevent garbage collection
        self.image_label.pack()

        # Load the image
        image = Image.open("logo.png")
        # Resize the image
        width, height = 150, 150
        image = image.resize((width, height), Image.ANTIALIAS)
        # Convert the Image object into a Tkinter PhotoImage object
        photo = ImageTk.PhotoImage(image)
        # Create a label and display the image
        self.image_label = tk.Label(self, image=photo,borderwidth=0)
        self.image_label.image = photo  # Keep a reference to the image to prevent garbage collection
        self.image_label.pack()
        self.image_label.place(x=50,y=15)

        # Frame to contain the buttons
        button_frame = tk.Frame(self)
        button_frame.pack(side=tk.TOP, fill=tk.X)
        button_frame.config(bg="white")


        self.generate_bill_button = tk.Button(button_frame, text="Generate Bill",font=("Verdana",14), command=self.show_generate_bill_form,bg="white",fg="#161F6D")
        self.generate_bill_button.pack(side=tk.LEFT)
        self.generate_bill_button.config(width=30,height=2,borderwidth=0)

        self.attendance_report_button = tk.Button(button_frame, text="Attendance Report", command=self.show_attendance_report_form,bg="#161F6D",fg="white",font=("Verdana",14))
        self.attendance_report_button.pack(side=tk.LEFT)
        self.attendance_report_button.config(width=30,height=2,borderwidth=0)

        self.change_configuration_button = tk.Button(button_frame, text="Change Configuration", command=self.show_change_configuration_form,bg="white",fg="#161F6D",font=("Verdana",14))
        self.change_configuration_button.pack(side=tk.LEFT)
        self.change_configuration_button.config(width=30,height=2,borderwidth=0)

        # Create a frame for the form with background color and border
        self.form_frame = tk.Frame(self, bg="#DDEDF4", padx=20, pady=20, relief=tk.RAISED, borderwidth=0)
        self.form_frame.place(y=100, relx=0.5, rely=0.5, anchor=tk.CENTER)

        # Select Attendance File
        self.attendance_label = tk.Label(self.form_frame, text="Select Attendance File:",bg="#DDEDF4",fg="#161F6D",font=("Franklin Gothic Demi",12))
        self.attendance_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.attendance_entry = tk.Entry(self.form_frame)
        self.attendance_entry.grid(row=0, column=1, padx=10, pady=10)
        self.attendance_entry.place(x=200,y=15,height=25,width=230)

        self.browse_button = tk.Button(self.form_frame, text="Browse",bg="#161F6D",fg="white",font=("Franklin Gothic Demi",10),borderwidth=0,width=12,command=self.select_attendance_file)
        self.browse_button.grid(row=0, column=2, padx=10, pady=15)

        # Select Branch
        self.branch_label = tk.Label(self.form_frame, text="Select Branch:",bg="#DDEDF4",fg="#161F6D",font=("Franklin Gothic Demi",12))
        self.branch_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.branch_combobox = ttk.Combobox(self.form_frame, values=["Computer", "Mechanical", "Electrical", "E&TC", "Civil"],width=34)
        self.branch_combobox.grid(row=1, column=1, padx=10, pady=10)

        # Select Semester
        self.semester_label = tk.Label(self.form_frame, text="Select Semester:",bg="#DDEDF4",fg="#161F6D",font=("Franklin Gothic Demi",12))
        self.semester_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.semester_combobox = ttk.Combobox(self.form_frame, values=["Odd", "Even"],width=34)
        self.semester_combobox.grid(row=2, column=1, padx=10, pady=10)

        # Select Session
        self.session_label = tk.Label(self.form_frame, text="Select Session:",bg="#DDEDF4",fg="#161F6D",font=("Franklin Gothic Demi",12))
        self.session_label.grid(row=3, column=0, padx=10, pady=10, sticky="w")
        self.session_combobox = ttk.Combobox(self.form_frame, values=[f"{year}-{year+1}" for year in range(2023, 2034)],width=34)
        self.session_combobox.grid(row=3, column=1, padx=10, pady=10)

        # Select Date Range
        self.date_range_label = tk.Label(self.form_frame, text="Select Date Range:",bg="#DDEDF4",fg="#161F6D",font=("Franklin Gothic Demi",12))
        self.date_range_label.grid(row=4, column=0, padx=10, pady=10, sticky="w")
        self.date_entry1 = DateEntry(self.form_frame)
        self.date_entry1.place(x=205,y=200,width=100,height=25)
        self.date_entry2 = DateEntry(self.form_frame)
        self.date_entry2.place(x=329,y=200,width=98,height=25)

        # Select Year
        self.year_label = tk.Label(self.form_frame, text="Select Year:",bg="#DDEDF4",fg="#161F6D",font=("Franklin Gothic Demi",12))
        self.year_label.grid(row=5, column=0, padx=10, pady=10, sticky="w")
        self.year_combobox = ttk.Combobox(self.form_frame, values=["First Year", "Second Year", "Third Year", "Final Year"],width=34)
        self.year_combobox.grid(row=5, column=1, padx=10, pady=10)

        # Generate Report Button
        
        self.generate_report_button = tk.Button(self, text="Generate Report",bg="#161F6D",fg="white",font=("Dubai Medium-Bold",14), command=self.call_generate_attendance_pdf)
        self.generate_report_button.pack(side=tk.LEFT)
        self.generate_report_button.config(height=2,borderwidth=0)
        self.generate_report_button.place(x=255,y=620,width=590)

    def select_attendance_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            self.attendance_entry.delete(0, tk.END)
            self.attendance_entry.insert(0, file_path)

    def show_change_configuration_form(self):
        self.clear_widgets()
        self.config(bg="white")
    
        # Load the image
        image = Image.open("bannerMain.png")

        # Resize the image
        width, height = 1100, 175
        image = image.resize((width, height), Image.ANTIALIAS)

        # Convert the Image object into a Tkinter PhotoImage object
        photo = ImageTk.PhotoImage(image)

        # Create a label and display the image
        self.image_label = tk.Label(self, image=photo,borderwidth=0)
        self.image_label.image = photo  # Keep a reference to the image to prevent garbage collection
        self.image_label.pack()

        # Load the image
        image = Image.open("logo.png")
        # Resize the image
        width, height = 150, 150
        image = image.resize((width, height), Image.ANTIALIAS)
        # Convert the Image object into a Tkinter PhotoImage object
        photo = ImageTk.PhotoImage(image)
        # Create a label and display the image
        self.image_label = tk.Label(self, image=photo,borderwidth=0)
        self.image_label.image = photo  # Keep a reference to the image to prevent garbage collection
        self.image_label.pack()
        self.image_label.place(x=50,y=15)
    
        # Frame to contain the buttons
        button_frame = tk.Frame(self)
        button_frame.pack(side=tk.TOP, fill=tk.X)
        button_frame.config(bg="white")
    
        self.generate_bill_button = tk.Button(button_frame, text="Generate Bill", font=("Verdana", 14), command=self.show_generate_bill_form, bg="white", fg="#161F6D")
        self.generate_bill_button.pack(side=tk.LEFT)
        self.generate_bill_button.config(width=30, height=2, borderwidth=0)
    
        self.attendance_report_button = tk.Button(button_frame, text="Attendance Report", command=self.show_attendance_report_form, bg="white", fg="#161F6D", font=("Verdana", 14))
        self.attendance_report_button.pack(side=tk.LEFT)
        self.attendance_report_button.config(width=30, height=2, borderwidth=0)
    
        self.change_configuration_button = tk.Button(button_frame, text="Change Configuration", command=self.show_change_configuration_form, bg="#161F6D", fg="white", font=("Verdana", 14))
        self.change_configuration_button.pack(side=tk.LEFT)
        self.change_configuration_button.config(width=30, height=2, borderwidth=0)
    
        # Create a frame for the form with background color and border
        self.form_frame = tk.Frame(self, bg="#DDEDF4", padx=20, pady=20, relief=tk.RAISED, borderwidth=0)
        self.form_frame.place(y=100, relx=0.5, rely=0.5, anchor=tk.CENTER)
    
        # Subject Code File
        self.subject_label = tk.Label(self.form_frame, text="Subject Code File:", bg="#DDEDF4", fg="#161F6D", font=("Franklin Gothic Demi", 12))
        self.subject_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.subject_entry = tk.Entry(self.form_frame)
        self.subject_entry.grid(row=0, column=1, padx=10, pady=5)
        self.upload_button = tk.Button(self.form_frame, text="Browse", bg="#161F6D", fg="white", font=("Franklin Gothic Demi", 10), borderwidth=0, width=10, command=self.upload_subject_file)
        self.upload_button.grid(row=0, column=2, padx=5, pady=5)
        self.remove_button = tk.Button(self.form_frame, text="Remove", bg="#161F6D", fg="white", font=("Franklin Gothic Demi", 10), borderwidth=0, width=10, command=lambda: self.subject_entry.delete(0, tk.END))
        self.remove_button.grid(row=0, column=3, padx=5, pady=5)
    
        # Theory, Other, and Limit
        self.theory_label = tk.Label(self.form_frame, text="Theory:", bg="#DDEDF4", fg="#161F6D", font=("Franklin Gothic Demi", 12))
        self.theory_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.theory_entry = tk.Entry(self.form_frame, width=20, validate="key")
        self.theory_entry.grid(row=1, column=1, padx=10, pady=5)
        self.theory_entry['validatecommand'] = (self.register(self.validate_numeric_input), '%S')
    
        self.other_label = tk.Label(self.form_frame, text="Other:", bg="#DDEDF4", fg="#161F6D", font=("Franklin Gothic Demi", 12))
        self.other_label.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.other_entry = tk.Entry(self.form_frame, width=20, validate="key")
        self.other_entry.grid(row=2, column=1, padx=10, pady=5)
        self.other_entry['validatecommand'] = (self.register(self.validate_numeric_input), '%S')
    
        self.limit_label = tk.Label(self.form_frame, text="Limit:", bg="#DDEDF4", fg="#161F6D", font=("Franklin Gothic Demi", 12))
        self.limit_label.grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.limit_entry = tk.Entry(self.form_frame, width=20, validate="key")
        self.limit_entry.grid(row=3, column=1, padx=10, pady=5)
        self.limit_entry['validatecommand'] = (self.register(self.validate_numeric_input), '%S')
    
        # Roll list file
        self.roll_entries = []
        for i, year in enumerate(["First", "Second", "Third", "Final"]):
            row = i + 4
            roll_list_label = tk.Label(self.form_frame, text=f"{year} Year Roll List:", bg="#DDEDF4", fg="#161F6D", font=("Franklin Gothic Demi", 12))
            roll_list_label.grid(row=row, column=0, padx=10, pady=5, sticky="w")
            roll_list_entry = tk.Entry(self.form_frame)
            roll_list_entry.grid(row=row, column=1, padx=10, pady=5)
            upload_button = tk.Button(self.form_frame, text="Browse", bg="#161F6D", fg="white", font=("Franklin Gothic Demi", 10), borderwidth=0, width=10, command=lambda e=roll_list_entry: self.upload_roll_file(e))
            upload_button.grid(row=row, column=2, padx=5, pady=5)
            remove_button = tk.Button(self.form_frame, text="Remove", bg="#161F6D", fg="white", font=("Franklin Gothic Demi", 10), borderwidth=0, width=10, command=lambda e=roll_list_entry: e.delete(0, tk.END))
            remove_button.grid(row=row, column=3, padx=5, pady=5)
            self.roll_entries.append(roll_list_entry)
    
        # Save and Reset Buttons
        self.save_button = tk.Button(self, text="Save", borderwidth=0, height=2, bg="#161F6D", fg="white", font=("Dubai Medium-Bold", 14), command=self.save_configuration)
        self.save_button.place(x=281, y=625, width=260)
    
        self.reset_button = tk.Button(self, text="Reset", borderwidth=0, height=2, bg="#161F6D", fg="white", font=("Dubai Medium-Bold", 14), command=self.reset_form)
        self.reset_button.place(x=556, y=625, width=260)
    
        # Load configuration data
        self.load_configuration()
        # Load configuration data
        self.load_configForm()
    
    def load_configuration(self):
    # Read the configuration file
      app_data_dir = self.get_app_data_directory()
      config_file_path = os.path.join(app_data_dir, "config.ini")
      if os.path.exists(config_file_path):
          config = configparser.ConfigParser()
          config.read(config_file_path)
          # Populate the form fields with the saved data
          if 'Paths' in config:
              self.subject_excel_path = config['Paths'].get('SubjectExcelPath', '')
              self.first_year_roll_list_path = config['Paths'].get('FirstYearRollListPath', '')
              self.second_year_roll_list_path = config['Paths'].get('SecondYearRollListPath', '')
              self.third_year_roll_list_path = config['Paths'].get('ThirdYearRollListPath', '')
              self.final_year_roll_list_path = config['Paths'].get('FinalYearRollListPath', '')
          if 'Amounts' in config:
              self.theory_amount = config['Amounts'].get('TheoryAmount', '')
              self.other_amount = config['Amounts'].get('OtherAmount', '')
              self.limit_amount = config['Amounts'].get('LimitAmount', '')
  
    def load_configForm(self):
        # Read the configuration file
        app_data_dir = self.get_app_data_directory()
        config_file_path = os.path.join(app_data_dir, "config.ini")
        if os.path.exists(config_file_path):
            config = configparser.ConfigParser()
            config.read(config_file_path)
            # Populate the form fields with the saved data
            if 'Paths' in config:
                self.subject_entry.insert(0, config['Paths'].get('SubjectExcelPath', ''))
                self.theory_entry.insert(0, config['Amounts'].get('TheoryAmount', ''))
                self.other_entry.insert(0, config['Amounts'].get('OtherAmount', ''))
                self.limit_entry.insert(0, config['Amounts'].get('LimitAmount', ''))
                self.roll_entries[0].insert(0, config['Paths'].get('FirstYearRollListPath', ''))
                self.roll_entries[1].insert(0, config['Paths'].get('SecondYearRollListPath', ''))
                self.roll_entries[2].insert(0, config['Paths'].get('ThirdYearRollListPath', ''))
                self.roll_entries[3].insert(0, config['Paths'].get('FinalYearRollListPath', ''))

                # # Check if the paths are already uploaded
                # subject_path = config['Paths'].get('SubjectExcelPath', '')
                # theory_path = config['Paths'].get('FirstYearRollListPath', '')
                # other_path = config['Paths'].get('SecondYearRollListPath', '')
                # limit_path = config['Paths'].get('ThirdYearRollListPath', '')
                # final_path = config['Paths'].get('FinalYearRollListPath', '')

                # # Set text accordingly
                # self.subject_entry.insert(0, "Already Uploaded" if os.path.exists(subject_path) else subject_path)
                # self.theory_entry.insert(0, "Already Uploaded" if os.path.exists(theory_path) else theory_path)
                # self.other_entry.insert(0, "Already Uploaded" if os.path.exists(other_path) else other_path)
                # self.limit_entry.insert(0, "Already Uploaded" if os.path.exists(limit_path) else limit_path)
                # self.roll_entries[0].insert(0, "Already Uploaded" if os.path.exists(final_path) else final_path)
                # self.roll_entries[1].insert(0, "Already Uploaded" if os.path.exists(final_path) else final_path)
                # self.roll_entries[2].insert(0, "Already Uploaded" if os.path.exists(final_path) else final_path)
                # self.roll_entries[3].insert(0, "Already Uploaded" if os.path.exists(final_path) else final_path)
    

    def validate_numeric_input(self, char):
        return char.isdigit() or char == '\b'

    def upload_subject_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            saved_path = self.save_file_securely(file_path)
            self.subject_entry.config(state="normal")
            self.subject_entry.delete(0, tk.END)
            self.subject_entry.insert(0, saved_path)
            self.subject_entry.config(state="readonly")

    def upload_roll_file(self, entry):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            saved_path = self.save_file_securely(file_path)
            entry.config(state="normal")
            entry.delete(0, tk.END)
            entry.insert(0, saved_path)
            entry.config(state="readonly")

    def save_file_securely(self, file_path):
        app_data_dir = self.get_app_data_directory()
        if not os.path.exists(app_data_dir):
            os.makedirs(app_data_dir)
        file_name = os.path.basename(file_path)
        saved_path = os.path.join(app_data_dir, file_name)
        shutil.copy(file_path, saved_path)
        return saved_path

    def get_app_data_directory(self):
        if os.name == 'posix':
            app_data_dir = os.path.expanduser("~/.config/your_app_name")
        elif os.name == 'nt':
            app_data_dir = os.path.join(os.environ['APPDATA'], 'your_app_name')
        else:
            app_data_dir = '.'
        os.makedirs(app_data_dir, exist_ok=True)
        return app_data_dir

    def save_configuration(self):
        subject_excel_path = self.subject_entry.get()
        first_year_roll_list_path = self.roll_entries[0].get()
        second_year_roll_list_path = self.roll_entries[1].get()
        third_year_roll_list_path = self.roll_entries[2].get()
        final_year_roll_list_path = self.roll_entries[3].get()
        theory_amount = self.theory_entry.get()
        other_amount = self.other_entry.get()
        limit_amount = self.limit_entry.get()

        config = configparser.ConfigParser()
        config['Paths'] = {
            'SubjectExcelPath': subject_excel_path,
            'FirstYearRollListPath': first_year_roll_list_path,
            'SecondYearRollListPath': second_year_roll_list_path,
            'ThirdYearRollListPath': third_year_roll_list_path,
            'FinalYearRollListPath': final_year_roll_list_path
        }
        config['Amounts'] = {
            'TheoryAmount': theory_amount,
            'OtherAmount': other_amount,
            'LimitAmount': limit_amount
        }

        app_data_dir = self.get_app_data_directory()
        config_file_path = os.path.join(app_data_dir, "config.ini")
        with open(config_file_path, 'w') as configfile:
            config.write(configfile)
        print("Configuration saved successfully.")

    def reset_form(self):
        self.subject_entry.config(state="normal")
        self.subject_entry.delete(0, tk.END)
        self.subject_entry.config(state="normal")

        for entry in self.roll_entries:
            entry.config(state="normal")
            entry.delete(0, tk.END)
            entry.config(state="normal")

        self.theory_entry.delete(0, tk.END)
        self.other_entry.delete(0, tk.END)
        self.limit_entry.delete(0, tk.END)

    def clear_widgets(self):
        for widget in self.winfo_children():
            widget.destroy()

    def call_generate_bill(self):
        print("function loadesd")
        self.load_configuration()

        file_path = self.file_entry.get()
        branch = self.branch_combobox.get()
        subject_file_path = self.subject_excel_path
        faculty_name = self.faculty_var.get()
        first_year_roll_list = self.first_year_roll_list_path
        second_year_roll_list = self.second_year_roll_list_path
        third_year_roll_list = self.third_year_roll_list_path
        final_year_roll_list = self.final_year_roll_list_path
        selected_month = self.month_var.get()
        # max_amount = self.limit_amount
        theory_amount = self.theory_amount
        other_amount = self.other_amount
        print(theory_amount, other_amount)
        # print(subject_file_path, first_year_roll_list)

        

        if not file_path or not branch or not subject_file_path or not faculty_name or not first_year_roll_list or not second_year_roll_list or not third_year_roll_list or not final_year_roll_list or not selected_month:
            messagebox.showerror("Missing Information", "Please select a file and fill in all the required details.")
            return
        
        try:
            generate_bill.generate_faculty_bill(file_path, branch, subject_file_path, faculty_name, first_year_roll_list, second_year_roll_list, third_year_roll_list, final_year_roll_list, selected_month, theory_amount, other_amount)
            ## calling form clearing function
            self.clear_bill_form()

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def clear_bill_form(self):
        # Clear the file entry
        self.file_entry.delete(0, tk.END)

        # Reset the faculty dropdown
        self.faculty_var.set('')
        self.faculty_dropdown['menu'].delete(0, 'end')

        # Clear the selected month
        for widget in self.month_checkboxes_frame.winfo_children():
            widget.destroy()
        self.month_var.set('')

        # Reset the branch combobox
        self.branch_combobox.set('')

    def call_generate_attendance_pdf(self):
        self.load_configuration()
        attendance_file_path = self.attendance_entry.get()
        branch = self.branch_combobox.get()
        semester = self.semester_combobox.get()
        session = self.session_combobox.get()
        start_date = self.date_entry1.get_date().strftime("%Y-%m-%d")
        end_date = self.date_entry2.get_date().strftime("%Y-%m-%d")
        year = self.year_combobox.get()
        subject_file_path = self.subject_excel_path

        if not attendance_file_path or not branch  or not semester or not session or not start_date or not end_date or not year:
           messagebox.showerror("Missing Information", "Please fill in all the required details.")
           return

        if not subject_file_path:
            messagebox.showerror("Please upload the Subject Code excel in change configuration")
            
        if year == "First Year":
            roll_list_path = self.first_year_roll_list_path
        elif year == "Second Year":
            roll_list_path = self.second_year_roll_list_path
        elif year == "Third Year":
            roll_list_path = self.third_year_roll_list_path
        elif year == "Final Year":
            roll_list_path = self.final_year_roll_list_path
        else:
            messagebox.showerror("Invalid Year", "Please select a valid year or check roll list of that  year is uploaded in configuration.")
            return
        
        try:
            generate_attendance.generate_attendance_pdf(roll_list_path, attendance_file_path, subject_file_path, branch, semester, session, start_date, end_date, year)
            self.clear_attendance_form()
        except Exception as e:
            print("Error", str(e))
    
    def clear_attendance_form(self):
        # Clear the attendance entry
        self.attendance_entry.delete(0, tk.END)

        # Reset the branch combobox
        self.branch_combobox.set('')

        # Reset the semester combobox
        self.semester_combobox.set('')

        # Reset the session combobox
        self.session_combobox.set('')

        # Reset the year combobox
        self.year_combobox.set('')

if __name__ == "__main__":
    start_window = StartWindow()
    start_window.mainloop()
