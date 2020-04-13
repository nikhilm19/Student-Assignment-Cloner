
from docx import Document

from docx2pdf import convert


import argparse

parser = argparse.ArgumentParser(description='Create mass copies of an Assignment')

parser.add_argument('--student_list_file', default=r'student-list.txt',
                   help='A txt file containing list of students')
parser.add_argument('--student_roll_file', default=r'student-roll.txt',
                   help='A txt file containing list of student roll nos')
parser.add_argument('--assignment_file', default=r'assignment.docx',
                   help='A docx file containing assignment')

parser.add_argument('--output_assignment_folder', default=r'assignments',
                   help='A folder to contain output copies')
args = parser.parse_args()


f=open(args.student_list_file,"r") # TXT file containing student names
lines= f.readlines()
f.close()

f=open(args.student_roll_file,"r") # TXT file containing roll numbers
roll_nos=f.readlines()
f.close()




document = Document(args.assignment_file)
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
	

	document.save(args.output_assignment_folder+"/"+roll+".docx")

	
convert(args.output_assignment_folder)