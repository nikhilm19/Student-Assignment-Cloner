
from docx import Document

from docx2pdf import convert


import argparse

parser = argparse.ArgumentParser(description='Create mass copies of an Assignment')

parser.add_argument('student_list_file', default='student-list.txt',
                   help='A txt file containing list of students')
parser.add_argument('assignment_file', default='assignment.docx',
                   help='A docx file containing assignment')

parser.add_argument('output_assignment_folder', default='assignments',
                   help='A folder to contain output copies')
args = parser.parse_args()


f=open(args.student_list_file,"r") # TXT file containing student names

lines= f.readlines()

roll_nos=[
160420107001,
160420107002,
160420107003,
160420107004,
160420107005,
160420107007,
160420107009,
160420107010,
160420107011,
160420107012,
160420107013,
160420107014,
160420107016,
160420107017,
160420107018,
160420107019,
160420107020,
160420107021,
160420107022,
160420107023,
160420107024,
160420107025,
160420107026,
160420107027,
160420107028,
160420107029,
160420107030,
160420107031,
160420107032,
160420107033,
160420107034,
160420107035,
160420107036,
160420107037,
160420107038,
160420107039,
160420107040,
160420107041,
160420107042,
160420107043,
160420107044,
160420107045,
160420107046,
160420107048,
160420107049,
160420107050,
160420107051,
160420107052,
160420107053,
160420107054,
160420107055,
160420107056,
160420107057,
160420107058,
160420107059,
160420107060,
160420107061,
160420107064,
160420107065,
160420107068,
160420107069,
160420107070,
160424107001,
170423107001,
170423107002,
170423107004,
170423107005,
170423107006,
170423107007,
170423107008,
170423107009,
170423107010,
170423107011,
170423107012,
170423107013,
170423107014,
170423107015
]




document = Document(args.assignment-file)
section = document.sections[0]
header = section.header
paragraph = header.paragraphs[1]
total_students = len(roll_nos)

for i in range(0,total_students):

	name=lines[i]
	roll=str(roll_nos[i])
	name= name.rstrip("\n")
	roll= roll.rstrip("\n")
	
	paragraph.text="Name:- "+name+"\t\t Enrollment No:-"+roll
	

	document.save(args.output-assignment-folder+"/"+roll+".docx")

	
convert("<path_to_your_assignment_folder>")