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


# Usage
transcript_generator = TranscriptGenerator()
start_roll = "0401CS01"
end_roll = "0401CS09"
transcript_generator.generate_marksheet(start_roll, end_roll)