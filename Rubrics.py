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
    contentGrades = pd.read_excel(file, 'NOTAS')
    dfGrades = pd.DataFrame(contentGrades)

    teacher = dfGrades.iloc[1][1]
    level = dfGrades.iloc[2][1] 
    semester = dfGrades.iloc[3][1]
    totalClasses = dfGrades.iloc[1][45]
    dfMocks = ""

    if ("expert" in level.lower() or "master" in level.lower()):
        contentMocks = pd.read_excel(file, 'MOCKS')
        dfMocks = pd.DataFrame(contentMocks)
    
    return dfGrades, dfMocks, teacher, level, semester, totalClasses

def CreateStudentList(parecer, dfGrades, dfMocks):
    student_grades = {}
    for i in range(5, dfGrades.index.stop):
        if(type(dfGrades.iloc[i][1]) == float):
            continue
        key = dfGrades.iloc[i][1]
        grades = dfGrades.iloc[i][2:]
        student_grades[key] = ValidateGrades(parecer, grades)
    
    if (type(dfMocks) != str):    
        for key in student_grades.keys():
            student_grades[key]["Mocks"] = {}
            for i in range(0, dfMocks.index.stop):
                if (dfMocks.iloc[i][0] != key):
                    continue
                for j in range(0,2):
                    student_grades[key]["Mocks"][f'Mock {j+1}'] = {}

                    student_grades[key]["Mocks"][f'Mock {j+1}']["Date"] = "" if (type(dfMocks.iloc[i+2+j][2]) == float) else dfMocks.iloc[i+2+j][2]
                    student_grades[key]["Mocks"][f'Mock {j+1}']["Test"] = "" if (type(dfMocks.iloc[i+2+j][3]) == float) else dfMocks.iloc[i+2+j][3]
                    student_grades[key]["Mocks"][f'Mock {j+1}']["Number"] = "" if (type(dfMocks.iloc[i+2+j][4]) == float) else dfMocks.iloc[i+2+j][4]
                    student_grades[key]["Mocks"][f'Mock {j+1}']["Reading and Use of English"] = round(dfMocks.iloc[i+2+j][20] * 100)
                    student_grades[key]["Mocks"][f'Mock {j+1}']["Listening"] = round(dfMocks.iloc[i+2+j][21] * 100)
                    student_grades[key]["Mocks"][f'Mock {j+1}']["Writing"] = round(dfMocks.iloc[i+2+j][22] * 100)
    
    return student_grades

def ValidateGrades(parecer, grades):
    columns = SelectColumns(parecer)
    listening = 0 if math.isnan(grades[columns[0]]) else round(grades[columns[0]],1)
    grammar = 0 if math.isnan(grades[columns[1]]) else round(grades[columns[1]],1)
    reading = 0 if math.isnan(grades[columns[2]]) else round(grades[columns[2]],1)
    writing = 0 if math.isnan(grades[columns[3]]) else round(grades[columns[3]],1)
    speaking = 0 if math.isnan(grades[columns[4]]) else round(grades[columns[4]],1)
    classperformance = 0 if math.isnan(grades[columns[5]]) else round(grades[columns[5]],1)
    parecergrade = 0 if math.isnan(grades[columns[6]]) else round(statistics.mean([speaking, listening, reading, writing, grammar,classperformance]),1)
    finalscore = 0 if math.isnan(grades[columns[6]]) else round(grades[columns[6]]/10, 1)
    attendance = 0 if math.isnan(grades[columns[7]]) else grades[columns[7]]
    comments = grades[columns[8]]

    return {"Grades": {"Listening" : listening, "Grammar" : grammar, "Reading" : reading, "Writing" : writing, "Speaking" : speaking, "Class Performance" : classperformance, "Final Score" : finalscore, "Parecer Grade" : parecergrade }, 
            "Attendance": attendance, 
            "Comments" : comments}

def SelectColumns(parecer):
    if (parecer == 1):
        columns = (2, 5, 8, 11, 14, 17, 42, 43, 44)
    else:
        columns = (20, 23, 26, 29, 32, 35, 42, 43, 45)
    return columns

def RenderTemplate(parecer, teacher, level, semester, totalClasses, folder_path, student_dict):

    loader = FileSystemLoader('templates')
    env = Environment(loader=loader)

    if ("expert" in level.lower() or "master" in level.lower()):
        if (parecer == 1):
            template = env.get_template('parecer1expert.html')
        else:
            template = env.get_template('parecer2expert.html')
    else:
        if (parecer == 1):
            template = env.get_template('parecer1.html')
        else :
            template = env.get_template('parecer2.html')

    for student in student_dict.keys():
        
        grammar = student_dict[student]["Grades"]["Grammar"]
        reading = student_dict[student]["Grades"]["Reading"]
        writing = student_dict[student]["Grades"]["Writing"]
        speaking = student_dict[student]["Grades"]["Speaking"]
        listening = student_dict[student]["Grades"]["Listening"]
        classperformance = student_dict[student]["Grades"]["Class Performance"]
        parecergrade = student_dict[student]["Grades"]["Parecer Grade"]
        finalscore = student_dict[student]["Grades"]["Final Score"]
        attendance = round((student_dict[student]["Attendance"]/totalClasses) * 100)
        comments = student_dict[student]["Comments"]
        mockTable = ""

        if (parecer == 1):
            gradeTable = getGradeTable(parecergrade)
        if (parecer == 2):
            gradeTable = getGradeTable(finalscore)

        if("Mocks" in student_dict[student]):
            mockTable = getMockTable(student_dict[student]["Mocks"])

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
                                attendance = attendance, 
                                comments = comments,
                                gradeTable = gradeTable,
                                mockTable = mockTable)

        file_path = f"{folder_path}/{student} - Parecer {parecer}"
        weasyprint.HTML(string = render, base_url = folder_path, encoding = 'UTF-8').write_pdf(f"{file_path}.pdf")
        print(f"{student} Parecer {parecer}.pdf successfully created at {folder_path}")

def getGradeTable(grade):

    distinction = '<td style="background-color: #FE3C00;"><span style="font-size: 10pt;">Distinction</span><br/><span style="font-size: 8pt;">Distinção</span></td>\n' if grade >= 9.5 else '<td><span style="font-size: 10pt;">Distinction</span><br/><span style="font-size: 8pt;">Distinção</span></td>\n'
    merit = '<td style="background-color: #FE3C00;"><span style="font-size: 10pt;">Merit</span><br/><span style="font-size: 8pt;">Mérito</span></td>\n' if (grade < 9.5 and grade >= 9) else '<td><span style="font-size: 10pt;">Merit</span><br/><span style="font-size: 8pt;">Mérito</span></td>\n'
    vgood = '<td style="background-color: #FE3C00;"><span style="font-size: 10pt;">Very Good</span><br/><span style="font-size: 8pt;">Muito Bom</span></td>\n' if (grade < 9 and grade >= 8) else '<td><span style="font-size: 10pt;">Very Good</span><br/><span style="font-size: 8pt;">Muito Bom</span></td>\n'
    good = '<td style="background-color: #FE3C00;"><span style="font-size: 10pt;">Good</span><br/><span style="font-size: 8pt;">Bom</span></td>\n' if (grade < 8 and grade >= 7) else '<td><span style="font-size: 10pt;">Good</span><br/><span style="font-size: 8pt;">Bom</span></td>\n'
    average = '<td style="background-color: #FE3C00;"><span style="font-size: 10pt;">Average</span><br/><span style="font-size: 8pt;">Regular</span></td>\n' if (grade < 7 and grade >= 6) else '<td><span style="font-size: 10pt;">Average</span><br/><span style="font-size: 8pt;">Regular</span></td>\n'
    fail = '<td style="background-color: #FE3C00;"><span style="font-size: 10pt;">Below Average</span><br/><span style="font-size: 8pt;">Abaixo da média</span></td>\n' if (grade < 6) else '<td><span style="font-size: 10pt;">Below Average</span><br/><span style="font-size: 8pt;">Abaixo da média</span></td>\n'

    tableContent = '<table class="concept">\n'
    tableContent += '<tr class="dark">\n'
    tableContent += '<th><span style="font-size: 10pt;">Grade</span><br/><span style="font-size: 8pt;">Nota</span></th>\n'
    tableContent += '<th><span style="font-size: 10pt;">Result</span><br/><span style="font-size: 8pt;">Resultado</span></th>\n'
    tableContent += '</tr>\n'
    tableContent += '<tr>\n'
    tableContent += '<td><span style="font-size: 10pt; white-space: nowrap;">9.5 a 10</span></td>\n'
    tableContent += distinction                
    tableContent += '</tr>\n'
    tableContent += '<tr class="dark">\n'
    tableContent += '<td><span style="font-size: 10pt;">9 a 9.4</span></td>\n'
    tableContent += merit                
    tableContent += '</tr>\n'
    tableContent += '<tr>\n'
    tableContent += '<td><span style="font-size: 10pt;">8 a 8.9</span></td>\n'
    tableContent += vgood
    tableContent += '</tr>\n'
    tableContent += '<tr class="dark">\n'
    tableContent += '<td><span style="font-size: 10pt;">7 a 7.9</span></td>\n'
    tableContent += good
    tableContent += '</tr>\n'
    tableContent += '<tr>\n'
    tableContent += '<td><span style="font-size: 10pt;">6 a 6.9</span></td>\n'
    tableContent += average
    tableContent += '</tr>\n'
    tableContent += '<tr class="dark">\n'
    tableContent += '<td><span style="font-size: 10pt;">Até 5.9</span></td>\n'
    tableContent += fail
    tableContent += '</tr>\n'
    tableContent += '</table>\n'

    return tableContent


def getMockTable(mocks):
    tableContent = '<table class="mocktable">\n'
    tableContent += '<tr>\n'
    tableContent += '<th style="font-size: 10pt;">Data</th>\n'
    tableContent += '<th style="font-size: 10pt;">Mock</th>\n'
    tableContent += '<th style="font-size: 10pt;">Número</th>\n'
    tableContent += '<th style="font-size: 10pt;">Reading and Use of English</th>\n'
    tableContent += '<th style="font-size: 10pt;">Writing</th>\n'
    tableContent += '<th style="font-size: 10pt;">Listening</th>\n'
    tableContent += '</tr>\n'

    for mock in mocks:
        tableContent += '<tr>\n'
        tableContent += '<td style="font-size: 10pt;">{}</td>\n'.format(mocks[mock]["Date"])
        tableContent += '<td style="font-size: 10pt;">{}</td>\n'.format(mocks[mock]["Test"])
        tableContent += '<td style="font-size: 10pt;">{}</td>\n'.format(mocks[mock]["Number"])
        tableContent += '<td class={}>{}%</td>\n'.format(getColor(mocks[mock]["Test"], mocks[mock]["Reading and Use of English"]),mocks[mock]["Reading and Use of English"])
        tableContent += '<td class={}>{}%</td>\n'.format(getColor(mocks[mock]["Test"], mocks[mock]["Listening"]),mocks[mock]["Listening"])
        tableContent += '<td class={}>{}%</td>\n'.format(getColor(mocks[mock]["Test"], mocks[mock]["Writing"]),mocks[mock]["Writing"])
        tableContent += '</tr>\n'

    tableContent += '</table>'
    return tableContent

def getColor(test, grade):

    if (grade < 45):
            return "fail"
    else:
        if (test == "FCE" and grade >= 80):
                return "c1"
        elif (test =="FCE" and (grade >= 60 and grade < 80)):
            return "b2"
        elif (test == "FCE" and (grade > 45 and grade < 60)):
            return "b1"
        
        elif (test == "CAE" and grade >= 80):
            return "c2"
        elif (test =="CAE" and (grade >= 60 and grade < 80)):
            return "c1"
        elif (test =="CAE" and grade > 45 and grade < 60):
            return "b2"
        
        elif (test == "CPE" and grade >= 60):
            return "c2"
        elif (test == "CPE" and (grade > 45 and grade < 60)):
            return "c1"
        
        else:
            return "fail"
        

def ConvertTable():
    parecer, files = SelectFiles()
    for file_path in files:
        folder_path = CreateFolder(file_path)
        dfGrades, dfMocks, teacher, level, semester, totalClasses = ExtractDataFrame(file_path)
        student_grades = CreateStudentList(parecer, dfGrades, dfMocks)
        RenderTemplate(parecer, teacher, level, semester, totalClasses, folder_path, student_grades)

ConvertTable()