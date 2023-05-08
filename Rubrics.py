import statistics
import pandas as pd
import os
import os.path
from jinja2 import Environment, FileSystemLoader
import math
import weasyprint
import PySimpleGUI as sg
import sys

def SelectFiles():
    sg.theme('Dark Grey 13')

    layout = [[sg.T("")],
            [sg.Text("Select the files: "), 
            sg.Input(), 
            sg.FilesBrowse(key="File Paths")],
            [sg.Radio('Parecer 1', "RADIO1", default=True, key="Parecer")],
            [sg.Radio('Parecer 2', "RADIO1", default=False)],
            [sg.Button("Submit")]]

    window = sg.Window('File Selection', layout)
    event, values = window.read()
    files = values["File Paths"].split(";")
    parecer = 1 if values["Parecer"] else 2

    if event == sg.WIN_CLOSED or event=="Exit":
        window.close()
        sys.exit()
    elif event == "Submit":
        window.close()
        if files[0] != '':
            return parecer, files
        else:
            sys.exit()

def CreateFolder(file_path):
    path = os.getcwd()
    file_name = os.path.split(file_path)[1].replace("_", " ")
    file_name = os.path.splitext(file_name)[0]
    folder_path = os.path.join(path, file_name)
    if(not os.path.exists(folder_path)):
        os.makedirs(folder_path)
    return folder_path

def ExtractDataFrame(file_path):
    file = pd.ExcelFile(file_path)
    content = pd.read_excel(file, 'NOTAS')
    df = pd.DataFrame(content)
    return df

def CreateStudentList(df):
    teacher = df.iloc[1][1]
    level = df.iloc[2][1]
    semester = df.iloc[3][1]
    student_grade_list = []
    for i in range(5, df.index.stop):
        if(type(df.iloc[i][1]) == float):
            continue
        student_grade_list.append(df.iloc[i])
    return teacher, level, semester, student_grade_list

def CreateStudentDict(parecer, student_list):
    student_grades_dict = {}
    for student in student_list:
        student_grades_dict[str(student[1])] = ValidateGrades(parecer, student)
    return student_grades_dict

def ValidateGrades(parecer, student):
    columns = SelectColumns(parecer)
    listening = 0 if math.isnan(student[columns[0]]) else round(student[columns[0]],1)
    grammar = 0 if math.isnan(student[columns[1]]) else round(student[columns[1]],1)
    reading = 0 if math.isnan(student[columns[2]]) else round(student[columns[2]],1)
    writing = 0 if math.isnan(student[columns[3]]) else round(student[columns[3]],1)
    speaking = 0 if math.isnan(student[columns[4]]) else round(student[columns[4]],1)
    classperformance = 0 if math.isnan(student[columns[5]]) else round(student[columns[5]],1)
    parecergrade = 0 if math.isnan(student[columns[6]]) else round(statistics.mean([speaking, listening, reading, writing, grammar,classperformance]) * 10)
    finalscore = 0 if math.isnan(student[columns[6]]) else int(round(student[columns[6]], 0))
    comments = student[columns[7]]

    return {"Listening" : listening, "Grammar" : grammar, "Reading" : reading, "Writing" : writing, "Speaking" : speaking, "Class Performance" : classperformance, "Final Score" : finalscore, "Parecer Grade" : parecergrade, "Comments" : comments}

def SelectColumns(parecer):
    if (parecer == 1):
        columns = (4, 7, 10, 13, 16, 19, 44, 46)
    else:
        columns = (22, 25, 28, 31, 34, 37, 44,47)
    return columns

def RenderTemplate(parecer, teacher, level, semester, folder_path, student_dict):

    loader = FileSystemLoader('templates')
    env = Environment(loader=loader)
    if (parecer == 1):
        template = env.get_template('parecer1.html')
    if (parecer == 2):
        template = env.get_template('parecer2.html')

    for student in student_dict:
        
        grammar = student_dict[student]["Grammar"]
        reading = student_dict[student]["Reading"]
        writing = student_dict[student]["Writing"]
        speaking = student_dict[student]["Speaking"]
        listening = student_dict[student]["Listening"]
        classperformance = student_dict[student]["Class Performance"]
        parecergrade = student_dict[student]["Parecer Grade"]
        finalscore = student_dict[student]["Final Score"]
        comments = student_dict[student]["Comments"]

        render = template.render(teacher = teacher,
                                level = level, 
                                semester = semester,
                                student = student, 
                                grammar = grammar,
                                reading = reading,
                                listening = listening,
                                writing = writing,
                                speaking = speaking,
                                classPerformance = classperformance,
                                parecerGrade = parecergrade,
                                finalScore = finalscore, 
                                comments = comments)

        file_path = f"{folder_path}/{student} - Parecer {parecer}"
        weasyprint.HTML(string = render, base_url = folder_path, encoding = 'UTF-8').write_pdf(f"{file_path}.pdf")
        print(f"{student} Parecer {parecer}.pdf successfully created at {folder_path}")

def ConvertTable():
    parecer, files = SelectFiles()
    for file_path in files:
        folder_path = CreateFolder(file_path)
        df = ExtractDataFrame(file_path)
        teacher, level, semester, student_grade_list = CreateStudentList(df)
        student_grades_dict = CreateStudentDict(parecer, student_grade_list)
        RenderTemplate(parecer, teacher, level, semester, folder_path, student_grades_dict)

ConvertTable()