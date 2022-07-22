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
    student_grade_list = []
    for i in range(5, df.index.stop):
        if(type(df.iloc[i][1]) == float):
            continue
        student_grade_list.append(df.iloc[i])
    return teacher, level, student_grade_list

def CreateStudentDict(parecer, student_list):
    student_grades_dict = {}
    for student in student_list:
        student_grades_dict[str(student[1])] = ValidateGrades(parecer, student)
    return student_grades_dict

def ValidateGrades(parecer, student):
    columns = SelectColumns(parecer)
    speaking = 0 if math.isnan(student[columns[0]]) else round(student[columns[0]],1)
    listening = 0 if math.isnan(student[columns[1]]) else round(student[columns[1]],1)
    reading = 0 if math.isnan(student[columns[2]]) else round(student[columns[2]],1)
    writing = 0 if math.isnan(student[columns[3]]) else round(student[columns[3]],1)
    grammar = 0 if math.isnan(student[columns[4]]) else round(student[columns[4]],1)
    classperformance = 0 if math.isnan(student[columns[5]]) else round(student[columns[5]],1)
    parecergrade = 0 if math.isnan(student[columns[6]]) else round(statistics.mean([speaking, listening, reading, writing, grammar]) * 10)
    finalscore = 0 if math.isnan(student[columns[6]]) else int(round(student[columns[6]], 0))
    comments = student[columns[7]]

    return {"Speaking" : speaking, "Listening" : listening, "Reading" : reading, "Writing" : writing, "Grammar" : grammar, "Class Performance" : classperformance, "Final Score" : finalscore, "Parecer Grade" : parecergrade, "Comments" : comments}

def SelectColumns(parecer):
    if (parecer == 1):
        columns = (4, 7, 10, 13, 16, 19, 44, 46)
    else:
        columns = (22, 25, 28, 31, 34, 37, 44,47)
    return columns

def RenderTemplate(parecer, teacher, level, folder_path, student_dict):

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
                                student = student, 
                                grammar = grammar, grammarfeedback = SelectGrammarFeedback(grammar),
                                reading = reading, readingfeedback = SelectReadingFeedback(reading),
                                listening = listening, listeningfeedback = SelectListeningFeedback(listening), 
                                writing = writing, writingfeedback = SelectWritingFeedback(writing),
                                speaking = speaking, speakingfeedback = SelectSpeakingFeedback(speaking),
                                classPerformance = classperformance, globalperformancefeedback = SelectGlobalPerformanceFeedback(classperformance),
                                parecerGrade = parecergrade,
                                finalScore = finalscore, comments = comments)

        file_path = f"{folder_path}/{student} - Parecer {parecer}"
        weasyprint.HTML(string = render, base_url = folder_path, encoding = 'UTF-8').write_pdf(f"{file_path}.pdf")
        print(f"{student} Parecer {parecer}.pdf successfully created at {folder_path}")

def SelectGrammarFeedback(grade):
    feedback = ""
    if grade > 9:
        feedback = """Apresenta facilidade com os aspectos gramaticais do idioma, utiliza estruturas gramaticais aprendidas previamente e coloca novas estruturas em prática durante a aula e nos temas de casa. Ótimo trabalho!"""
    elif grade > 8:
        feedback = """Utiliza estruturas gramaticais aprendidas previamente com boa precisão e coloca novas estruturas em prática durante a aula e nos temas de casa. Sempre que necessário, revise as estruturas anteriores e tire suas dúvidas com seu professor. Busque usar, sempre que possível, as estruturas que aprendeu em sua fala e escrita. Mantenha o bom trabalho!"""
    elif grade > 7:
        feedback = """Utiliza estruturas gramaticais aprendidas previamente com boa desenvoltura e com alguma frequência coloca novas estruturas em prática durante a aula e nos temas de casa. Sempre que necessário, revise as estruturas anteriores e tire suas dúvidas com seu professor. Busque usar, sempre que possível, as estruturas e vocabulário que aprendeu em sua fala e escrita. Você está no caminho certo."""
    elif grade > 6:
        feedback = """Um novo nível sempre apresenta novos desafios - e você está no caminho certo! Busque sempre colocar em prática as estruturas e vocabulário que aprendemos em aula, e tire dúvidas sempre que possível. Quando necessário, revise as estruturas que aprendemos anteriormente. A prática faz a perfeição, portanto, mantenha o foco, realize suas atividades e deveres de casa e busque maiores resultados e desafios!"""
    else:
        feedback = """Seu desempenho nesta habilidade inspira cuidados. Busque sempre colocar em prática as estruturas e vocabulário que aprendemos em aula, e tire dúvidas sempre que possível. Quando necessário, revise as estruturas que aprendemos anteriormente. A prática faz a perfeição, portanto, mantenha o foco, realize suas atividades e deveres de casa e busque maiores resultados e desafios!"""
    return feedback

def SelectReadingFeedback(grade):
    feedback = ""
    if grade > 9:
        feedback = """Apresenta ótima compreensão escrita, faz bom uso de estratégias ao realizar exercícios de reading e consegue entender o contexto geral dos textos trabalhados. Facilidade em captar os detalhes específicos pedidos. Ótima desenvoltura!"""
    elif grade > 8:
        feedback = """Apresenta ótima compreensão escrita, faz bom uso de estratégias ao realizar exercícios de reading e consegue entender o contexto geral dos textos trabalhados, com pequenos deslizes no entendimento de detalhes específicos. Continue o bom trabalho, e lembre-se sempre de praticar essa habilidade cercando-se de conteúdo relevante em inglês. Leia sobre aquilo que lhe é interessante, e descubra coisas novas sobre aquilo que não é! Continue assim para progredir cada vez mais!"""
    elif grade > 7:
        feedback = """Apresenta boa compreensão escrita, faz uso de estratégias ao realizar exercícios de reading e consegue entender o contexto geral dos textos trabalhados. Apresenta alguma dificuldade em detalhes específicos. Desenvolva ainda mais essa habilidade lendo textos de gêneros variados, inclusive os que não leria habitualmente. Divirta-se e descubra coisas novas em inglês! Continue assim para progredir cada vez mais!"""
    elif grade > 6:
        feedback = """Compreende o contexto geral dos textos apresentados, e apresenta grande potencial de aprimoramento da habilidade. Estratégias para otimizar a leitura e captar detalhes específicos são a chave do sucesso - converse com o seu professor a respeito. Ler é um ótimo exercício para polir e expandir seu vocabulário; não tenha medo de revisar os textos utilizados em aula e buscar novas descobertas em livros e na internet. Aqui vão algumas sugestões: faça uma rápida leitura inicial para entender a ideia geral do texto com o qual está lidando. Leia sempre com cuidado a atividade proposta e suas questões. Cultive sempre o hábito da leitura, e divirta-se. Ler nos traz experiências únicas!"""
    else:
        feedback = """Seu desempenho nesta habilidade inspira cuidados. Estratégias para otimizar a leitura e captar detalhes específicos são a chave do sucesso - converse com o seu professor a respeito. Ler é um ótimo exercício para polir e expandir seu vocabulário; não tenha medo de revisar os textos utilizados em aula e buscar novas descobertas em livros e na internet. Aqui vão algumas sugestões: faça uma rápida leitura inicial para entender a ideia geral do texto com o qual está lidando. Leia sempre com cuidado a atividade proposta e suas questões. Cultive sempre o hábito da leitura, e divirta-se. Ler nos traz experiências únicas!"""
    return feedback

def SelectListeningFeedback(grade):
    feedback = ""
    if grade > 9:
        feedback = """Ótimo entendimento do que é dito e de todas as atividades envolvendo listening. Consegue entender o contexto geral nas passagens de áudio, bem como captar os detalhes específicos pedidos. Faz bom uso de estratégias ao realizar exercícios de listening. Continue assim!"""
    elif grade > 8:
        feedback = """Bom entendimento do que é dito e de todas as atividades envolvendo listening. Consegue entender o contexto geral nas passagens de áudio, bem como captar os detalhes específicos pedidos. Faz bom uso de estratégias ao realizar exercícios de listening. Procure assistir filmes/vídeos/séries e ouvir músicas em inglês com sotaques variados - isso pode ajudar muito em seu progresso nessa habilidade. Bom trabalho, busque aprimorar-se ainda mais!"""
    elif grade > 7:
        feedback = """Bom entendimento do que é dito e de atividades envolvendo listening. Consegue entender o contexto geral nas passagens de áudio, bem como captar os detalhes específicos pedidos com algum esforço. Faz bom uso de estratégias ao realizar exercícios de listening. Procure assistir filmes/vídeos/séries e ouvir músicas em inglês com sotaques variados - isso pode ajudar muito em seu progresso nessa habilidade. Bom trabalho, busque aprimorar-se ainda mais!"""
    elif grade > 6:
        feedback = """Sabe aquela música que não sai da sua cabeça? Ela pode ser uma grande aliada no desenvolvimento dessa habilidade! Como todas as outras, ela precisa de prática. Tente entender a ideia geral do que está sendo dito - ao invés de focar em palavra por palavra. Leia com cuidado as atividades antes de ouvir o áudio, e destaque palavras que podem ajudá-lo a resolver questões. Converse com o seu professor a respeito de estratégias para aprimorar sua habilidade de entender detalhes específicos. Em casa, assista aquele filme que você gosta, ouça aquela música que você quer muito aprender a cantar. Quem disse que aprender inglês tem que ser chato? Aprimorar o listening pode ser divertido!"""
    else:
        feedback = """Seu desempenho nesta habilidade inspira cuidados. Sabe aquela música que não sai da sua cabeça? Ela pode ser uma grande aliada no desenvolvimento dessa habilidade! Como todas as outras, ela precisa de prática. Tente entender a ideia geral do que está sendo dito - ao invés de focar em palavra por palavra. Leia com cuidado as atividades antes de ouvir o áudio, e destaque palavras que podem ajudá-lo a resolver questões. Converse com o seu professor a respeito de estratégias para aprimorar sua habilidade de entender detalhes específicos. Em casa, assista aquele filme que você gosta, ouça aquela música que você quer muito aprender a cantar. Quem disse que aprender inglês tem que ser chato? Aprimorar o listening pode ser divertido!"""
    return feedback

def SelectWritingFeedback(grade):
    feedback = ""
    if grade > 9:
        feedback = """Expressa claramente suas ideias e as conecta de forma coerente. Faz uso de estruturas pertinentes ao nível, cumpre as instruções e atinge os objetivos dos textos. Busque sempre diversificar seu vocabulário ainda mais! Excelente trabalho!"""
    elif grade > 8:
        feedback = """Expressa de maneira muito satisfatória suas ideias e as conecta de forma coerente. Faz bom uso de estruturas pertinentes ao nível, cumpre em sua maior parte as instruções e atinge os objetivos dos textos. Busque sempre diversificar seu vocabulário ainda mais, e não tenha medo de ser ambicioso em seus textos. Esperamos ansiosos por sua próxima aventura na escrita!"""
    elif grade > 7:
        feedback = """Expressa de maneira satisfatória suas ideias e as conecta de forma coerente. Faz uso de estruturas pertinentes ao nível, cumpre em sua maior parte as instruções e atinge os objetivos dos textos. Busque sempre diversificar seu vocabulário, utilizar as estruturas que aprendemos e ser ambicioso em seus textos. Continue desbravando os misteriosos mundos da escrita!"""
    elif grade > 6:
        feedback = """Grandes ideias são o início de todo bom texto - mas nem sempre é fácil colocá-las no papel. Para isso, tente organizar suas ideias e decidir onde quer chegar com seu trabalho, e tente sempre utilizar as estruturas e vocabulários que aprendemos em aula. Antes de entregar seu texto, verifique sua ortografia, pontuação, se conseguiu atingir todos os objetivos que a questão pede. Seu texto faz sentido? Diz aquilo que você queria dizer? A escrita, como qualquer habilidade, requer prática. Converse com o seu professor caso necessite de apoio."""
    else:
        feedback = """Seu desempenho nesta habilidade inspira cuidados. Tente organizar suas ideias e decidir onde quer chegar com seu trabalho, e tente sempre utilizar as estruturas e vocabulários que aprendemos em aula. Antes de entregar seu texto, verifique sua ortografia, pontuação, se conseguiu atingir todos os objetivos que a questão pede. Seu texto faz sentido? Diz aquilo que você queria dizer? A escrita, como qualquer habilidade, requer prática. Converse com o seu professor caso necessite de apoio."""
    return feedback

def SelectSpeakingFeedback(grade):
    feedback = ""
    if grade > 9:
        feedback = """Apresenta boa fluência e desenvoltura na fala, e faz uso adequado e constante do inglês em sala de aula. Possui boa pronúncia e entonação, e faz uso adequado de vocabulário e estruturas pertinentes ao nível. Continue assim!"""
    elif grade > 8:
        feedback = """Apresenta boa desenvoltura na fala, e faz uso adequado do inglês em sala de aula. Tente sempre colocar em prática o que aprendeu em aula, e não tenha medo de errar. Solte a língua! Buscar sempre mais conhecimento e interação é o caminho para atingir notas ainda melhores!"""
    elif grade > 7:
        feedback = """Apresenta desenvoltura satisfatória, com alguma fluência na fala. Tente sempre colocar em prática o que aprendeu em aula, e não tenha medo de errar. Dessa maneira, sua desenvoltura nessa habilidade ficará cada vez melhor! Solte a língua! Buscar sempre mais conhecimento e interação é o caminho para atingir notas ainda melhores!"""
    elif grade > 6:
        feedback = """Hey, estamos sentindo falta de ouvir a sua voz! Interagir em sala de aula, com seu professor e colegas, é a chave para um bom desenvolvimento da sua habilidade de fala. Tente utilizar as estruturas e vocabulário que aprendemos, e não tenha medo de errar - errar é uma experiência de aprendizado. Busque usar o inglês ao máximo, e não tenha receio de pedir ajuda ao professor. Recorrer ao português deve ser um recurso utilizado apenas em situações de tradução de palavras novas. Se tiver oportunidade, fale inglês fora da sala de aula também."""
    else:
        feedback = """Seu desempenho nesta habilidade inspira cuidados. Interagir em sala de aula, com seu professor e colegas, é a chave para um bom desenvolvimento da sua habilidade de fala. Tente utilizar as estruturas e vocabulário que aprendemos, e não tenha medo de errar - errar é uma experiência de aprendizado. Busque usar o inglês ao máximo, e não tenha receio de pedir ajuda ao professor. Recorrer ao português deve ser um recurso utilizado apenas em situações de tradução de palavras novas. Se tiver oportunidade, fale inglês fora da sala de aula também."""
    return feedback

def SelectGlobalPerformanceFeedback(grade):
    feedback = ""
    if grade > 9:
        feedback = """Muito participativo, tem boa interação com colegas e professor. É assíduo e solícito. Mantenha o ótimo desempenho!"""
    elif grade > 8:
        feedback = """Participativo, tem boa interação com colegas e professor. É assíduo e solícito, e busca tirar dúvidas e fazer deveres de casa. Busca estar envolvido nas discussões e brincadeiras feitas em aula, é pontual e comprometido com suas tarefas. Ótimo desempenho!"""
    elif grade > 7:
        feedback = """Participativo, tem boa interação com colegas e professor. Em geral, é assíduo e solícito, e busca tirar dúvidas e fazer deveres de casa. Busque sempre estar envolvido nas discussões e atividades feitas em aula, ser pontual e comprometido com suas tarefas. Você já está no caminho - estamos quase lá!"""
    elif grade > 6:
        feedback = """Participa com alguma frequência, e pode interagir cada vez mais com colegas e professor. Busque sempre ser assíduo, tirar suas dúvidas e fazer os deveres de casa. Envolva-se nas discussões e atividades feitas em aula, seja pontual e comprometido com o seu próprio desenvolvimento. A aula fica mais divertida quando participamos ativamente!"""
    else:
        feedback = """Seu desempenho nesta habilidade inspira cuidados. Pode interagir cada vez mais com colegas e professor. Busque sempre ser assíduo, tirar suas dúvidas e fazer os deveres de casa. Envolva-se nas discussões e atividades feitas em aula, seja pontual e comprometido com o seu próprio desenvolvimento. A aula fica mais divertida quando participamos ativamente!"""
    return feedback

def ConvertTable():
    parecer, files = SelectFiles()
    for file_path in files:
        folder_path = CreateFolder(file_path)
        df = ExtractDataFrame(file_path)
        teacher, level, student_grade_list = CreateStudentList(df)
        student_grades_dict = CreateStudentDict(parecer, student_grade_list)
        RenderTemplate(parecer, teacher, level, folder_path, student_grades_dict)

ConvertTable()