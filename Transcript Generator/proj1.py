import csv
import os
from fpdf import FPDF
'''The FPDF library is a PHP class that allows you to generate PDF files with pure PHP, 
without using the PDFlib library1. It is free to use and modify, and it has many 
features such as image support, colors, links, fonts, page compression, etc'''

from pywebio import *
from pywebio.input import *
from pywebio.output import *
'''pywebio: A Python package that allows you to write interactive web applications or 
browser-based GUI applications without the need to have knowledge of HTML and JS'''


class TranscriptGenerator:
    def __init__(self):
        self.grades_dict = {"AA": 10,"AB": 9,"BB": 8,"BC": 7,"CC": 6,"CD": 5,"DD": 4,"DD*": 4,"F": 0,"F*": 0,"I": 0}
        self.stud_dict = {}
        self.courses_dict = {}
        self.heading = ["Sub. Code", "Subject Name", "L-T-P", "CRD", "GRD"]
        self.student_details = {'roll': "1901EE55", 'name': "Amelia Jonas",
                                'year': "2019", "programme": "Bachelor of Technology",
                                'course': "Computer science and engineering"}
        self.course_details = {
            "CS": "Computer Science and Engineering",
            "EE": "Electrical Engineering",
            "ME": "Mechanical Engineering"
        }

        # self.pdf = FPDF('L' ,'mm' , (800 , 830))

    def pre_computation(self):
        print("pre computation starting")
        with open("sample_input/subjects_master.csv", "r") as file:
            reader = csv.DictReader(file)
            for i in reader:
                self.courses_dict[i["subno"]] = {"subname": i["subname"], "ltp": i["ltp"], "crd": i["crd"]}

        with open("sample_input/names-roll.csv", "r") as file:
            reader = csv.DictReader(file)
            for i in reader:
                self.stud_dict[i["Roll"]] = {"Name": i["Name"]}

        with open("sample_input/grades.csv", "r") as file:
            reader = csv.DictReader(file)
            for i in reader:
                try:
                    self.stud_dict[i["Roll"]][i["Sem"]]
                except KeyError:
                    self.stud_dict[i["Roll"]][i["Sem"]] = {}
                self.stud_dict[i["Roll"]][i["Sem"]][i["SubCode"]] = {"Grade": i["Grade"].strip(), "Sub_Type": i["Sub_Type"]}
        print("Pre computation done")


    # Method to create a PDF consisting of LOGO AND NAME of IIT Patna
    def generate_pdf(self):
        self.pdf.add_page()
        start_index_x = 30
        self.pdf.set_font('Arial' , 'B' , 16)
        self.pdf.set_left_margin(20)
        self.pdf.set_right_margin(20)
        self.pdf.cell(0 , 700 , "" ,  1, 1)
        self.pdf.set_xy(20 , 10)
        self.pdf.cell(80, 80, "", 1,0,'C')
        self.pdf.set_xy(20 , 10)
        self.pdf.image('logo.png' ,self.pdf.get_x() + 5 , self.pdf.get_y() + 5 , 70 , 70 , "png", 'logo.png')
        self.pdf.set_xy(100 , 10)
        self.pdf.cell(630 , 80 , "" , 1 , 0 , 'C')
        self.pdf.set_xy(100 , 10)
        self.pdf.image('name.jpg', self.pdf.get_x() + 5, self.pdf.get_y() + 5 , 625, 70, "png", 'name.jpg' )
        self.pdf.set_xy(730 , 10)
        self.pdf.cell(80, 80, "", 1,0,'C')
        self.pdf.set_xy(735 , 10)
        self.pdf.image('logo.png' ,self.pdf.get_x() , self.pdf.get_y() + 5 , 70 , 70 , "png", 'logo.png')
        self.pdf.set_y(80)
        self.pdf.set_x(start_index_x)


    def make_description(self):
        self.pdf.set_y(self.pdf.get_y())
        self.pdf.set_x(250)
        self.pdf.cell(400, 16,"" ,1, 0, 'C')
        self.pdf.set_x(280)
        self.pdf.set_font("Arial", 'B', 16)
        self.pdf.cell(120, 8, f"Roll No:  {self.student_details['roll']}" , 0, 0)
        self.pdf.cell(120, 8, f"Name:  {self.student_details['name']}" , 0, 0)
        self.pdf.cell(120, 8, f"Year of admission:  {self.student_details['year']}" , 0, 1)
        self.pdf.set_x(280)
        self.pdf.cell(120, 8, f"Programme:  {self.student_details['programme']}" , 0, 0)
        self.pdf.cell(120, 8, f"Course:  {self.student_details['course']}" , 0, 1)
        self.pdf.set_y(self.pdf.get_y() + 10)


    def set_coordinates(self, x, y, sem):
        # print(int(sem/4) , (int(sem%3)-1))
        self.pdf.set_y(y + int((sem-1)/3)*150)
        self.pdf.set_x(x + (int((sem-1)%3))*240 + (int((sem-1)%3))*20)
        if (sem-1)%3 == 0 and sem!=1 : 
            self.make_line(self.pdf.get_y())
            # self.pdf.set_y(self.pdf.get_y() + 20)
    
    def make_line(self , y) :
        self.pdf.line(20, y, 810, y)


    # Method to calculate CPI
    def cpi_calc(self, grades , credits):  # function to calculate the cpi upto a particular semester 
        tot_sum = 0.0
        cred_sum = 0
        for i in range(len(grades)):
            tot_sum += grades[i]*credits[i]
        for c in credits:
            cred_sum += c
        return round(tot_sum/cred_sum, 2)  #return the rounded cpi upto 2 decimal
    

    # Displaying the Credits, SPI, and CPI of the students
    def overall_credits_cell(self, details) :
        self.pdf.set_font('Arial', 'B', 16)
        self.pdf.cell(200 , 10, f"Credits Taken: {details['credits']}    Credits Cleared: {details['credits']}  SPI: {details['spi']}   CPI: {details['cpi']}", 1, 2)
        self.pdf.set_font('Arial' , '', 16)


    def semester_name(self, name):
        self.pdf.set_font('Arial',"BU", 16)
        self.pdf.cell(30, 10, f"Semester {name}", 0, 2)
        self.pdf.set_font('Arial','',16)

    def create_cell(self, type, to, content):
        # print(type ,to , content)
        if type==1 :
            # print("this is type 1")
            self.pdf.cell(60, 10, str(content) , 1, to, 'C')
            return

        if type==2 :
            # print("this is type 2")
            self.pdf.cell(140, 10, str(content) , 1, to, 'C')
            return

        if type==3 :
            # print("this is type 3")
            self.pdf.cell(20, 10, str(content) , 1, to, 'C')
            return

        if type==4 :
            # print("this is type 4")
            self.pdf.cell(15, 10, str(content) , 1, to, 'C')
            return

        if type==5 :
            # print("this is type 5")
            self.pdf.cell(15, 10, str(content) , 1, to, 'C')
            return
        if type==6 :
            # print("this is type 6")
            self.pdf.cell(100 , 10, str(content), 1, to, 'C')


    def create_table(self, start_x, table_body):
        # print(headers)
        # semester_name(self.pdf,1)
        self.pdf.set_x(start_x)
        self.pdf.set_font("Arial", 'BU', 16)
        i = 1

        for h in self.heading :
                if i < 5:
                    self.create_cell(i , 0, h)
                else : 
                    self.create_cell(i, 1 , h)
                i = i+1

        self.pdf.set_x(start_x)
        self.pdf.set_font("Arial", '', 16)

        i = 1
        for row in table_body:
            for column in row :
                if i < 5:
                    self.create_cell(i , 0, column)
                else : 
                    self.create_cell(i, 1 , column)
                i = i+1
            self.pdf.set_x(start_x)
            i = 1


    # Setting the footer format of the pdf
    def footer(self) :
        self.pdf.set_y(self.pdf.get_y() + 80)
        y = self.pdf.get_y()
        self.pdf.set_x(30)
        self.pdf.set_font('Arial', 'B', 16)
        self.pdf.text(self.pdf.get_x() , self.pdf.get_y(), "Date of issue: ")
        self.pdf.line(75, self.pdf.get_y(), 150, self.pdf.get_y())
        self.pdf.set_y(y - 10)
        self.pdf.set_x(700)
        self.pdf.cell(80, 1, "", 'B', 2, 'C')
        self.pdf.cell(80, 10, "Assistant Registrar (Academic) ", 0, 0, 'C')
            

    def generate_marksheet(self, start_roll, end_roll):
        # Calling pre_computational method to get all the input data into respective dictionary
        self.pre_computation()

        if os.path.exists("transcriptsIITP") == False:
            os.makedirs("transcriptsIITP")
        prefix = start_roll[0:6]

        if start_roll[4:6] != end_roll[4:6]:
            print("The students must be from the same department")
            return

        if end_roll[0:6] != prefix:
            print("Please enter the correct range of roll numbers")
            return

        start = int(start_roll[6:])
        end = int(end_roll[6:])
        if start > end:
            print("Please input the correct range of roll numbers")
            return

        print(start, end)

        print("Your Marksheets are being generated......")
        not_present_roll_no = []

        for i in range(start , end+1):
            curr_roll = prefix + str(int(i/10)) + str(int(i%10))
            print(curr_roll)
            if curr_roll not in self.stud_dict :
                not_present_roll_no.append(curr_roll)
                continue 

            self.pdf = FPDF('L' ,'mm' , (800 , 830))

            # Calling generate_pdf method 
            self.generate_pdf()

            self.pdf.set_y(self.pdf.get_y() + 20)
            self.student_details['roll'] = curr_roll
            self.student_details['course'] = self.course_details[curr_roll[4:6]]

            # Calling make_description method 
            self.make_description()

            self.pdf.set_x(30)

            credits = [0,0,0,0,0,0,0,0]                #list to store the credits sum of a semester 
            total_credits = [0,0,0,0,0,0,0,0]          #list to store the credits sum till semester 
            spi = [0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0]    #list to store the spi of all the semesters 
            cpi = [0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0]    #list to store the cpi of all the semesters

            x_coordinate = self.pdf.get_x()
            y_coordinate = self.pdf.get_y()

            for j in range(1, 9):
                self.set_coordinates(x_coordinate , y_coordinate, j)
                creds = []
                grades = []
                l = 1

                try:
                    self.stud_dict[curr_roll][str(j)]
                except KeyError:
                    continue

                table_data = []

                for k in self.stud_dict[curr_roll][str(j)]: #k is the subject code 
                    temp_data_row = []
                    temp_data_row.append(k)
                    temp_data_row.append(self.courses_dict[k]['subname'])
                    temp_data_row.append(self.courses_dict[k]["ltp"])
                    temp_data_row.append(self.courses_dict[k]["crd"])
                    temp_data_row.append(self.grades_dict[self.stud_dict[curr_roll][str(j)][k]["Grade"]])
                    table_data.append(temp_data_row)
                    creds.append(int(self.courses_dict[k]["crd"]))
                    grades.append(self.grades_dict[self.stud_dict[curr_roll][str(j)][k]["Grade"]])
                    l += 1

                self.semester_name(j)
                self.create_table(self.pdf.get_x(), table_data)

                for c in creds:
                    credits[j-1]+=c
                spi[j-1] = self.cpi_calc(grades, creds)
                if j>1:
                    total_credits[j-1] = total_credits[j-2]+credits[j-1]
                    cpi[j-1] = self.cpi_calc(spi[:j],credits[:j])
                else:
                    total_credits[j-1]=credits[j-1]
                    cpi[j-1]=spi[j-1]
                details = {'credits' : credits[j-1] , 'spi' : spi[j-1], 'cpi' : cpi[j-1]}

                # self.pdf.set_y(self.pdf.get_y() + 5)
                self.overall_credits_cell(details)

            self.set_coordinates(x_coordinate , y_coordinate, 10)
            self.footer()
            self.pdf.output("transcriptsIITP/"+curr_roll+".pdf")

        print(not_present_roll_no)
        print("The required transcripts have been generated, please look in the transcriptsIITP folder for the same")
        for roll in not_present_roll_no :
            print(f"Roll no {roll} was not found in the list")


# Usage
# transcript_generator = TranscriptGenerator()
# start_roll = "0401CS01"
# end_roll = "0401CS09"
# transcript_generator.generate_marksheet(start_roll, end_roll)


def main() :
    req_action = "Generate Marksheets"
    while(req_action != "None") :

        req_action = actions(label = "Please select your action" ,
                    buttons=[{'label' : "Generate Marksheets",'value':"Generate Marksheets"} ,
                        {'label' : "I wish to exit",'value':"None"}
                    ])

        if (req_action == "Generate Marksheets"):
            start = input("Please enter the starting roll number", type = TEXT)
            end = input("Please enter the ending Roll numbetr to generate the report", TYPE=TEXT)
            transcript_generator = TranscriptGenerator()
            transcript_generator.generate_marksheet(start, end)
        else :
            put_success("Hope you liked the project")
            
start_server(main, port=3001)  