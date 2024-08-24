import cv2
import skimage as ski
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread.cell import Cell
import pytesseract
pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract.exe'
import time
import matplotlib.pyplot as plt
import numpy as np

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('cursinhox-47bc2e88124b.json', scope)
client = gspread.authorize(credentials)

## VARIÁVEIS
documentName = 'Cópia de Correção Simulado LF - 30/06 (respostas)'

## CLASSES
class Coordinates:
    def __init__(self, x, y):
        self.x = x
        self.y = y
    def __repr__(self):
        return repr((self.x, self.y))

class Question:
    def __init__(self, coord):
        self.coord = coord
        self.num = 0
        self.ans = ''
    def __repr__(self):
        return repr((self.coord, self.num, self.ans))

class AnsMap:
    def __init__(self, x, ans, section):
        self.x = x
        self.ans = ans
        self.section = section
    def __repr__(self):
        return repr((self.x, self.ans, self.section))

class QMap:
    def __init__(self, y, number, section):
        self.y = y
        self.number = number
        self.section = section
    def __repr__(self):
        return repr((self.y, self.number, self.section))

## PROGRAMA
answersMap = []
questionsMap = []

def preprocess_image(image):
    image_gray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)
    #len_img_header = len(image_gray) // 4
    #image_gray[:int(len_img_header)] = 255
    image_RGB = cv2.cvtColor(image_gray, cv2.COLOR_GRAY2RGB)
    return image_gray, image_RGB

def detect_elements(image_gray):
    blurred = cv2.medianBlur(image_gray, 21)
    canny = cv2.Canny(image_gray, 200, 300)
    
    detected_answers = cv2.HoughCircles(
        blurred, cv2.HOUGH_GRADIENT, 1, 
        minDist=20, param1=40, param2=15, 
        minRadius=10, maxRadius=20
    )
    
    detected_questions = cv2.HoughCircles(
        canny, cv2.HOUGH_GRADIENT, 1, 
        minDist=20, param1=40, param2=15, 
        minRadius=10, maxRadius=20
    )
    
    return detected_answers, detected_questions


def getJPEGItensInPath(path):
    files = os.listdir(path)
    result = []
    for el in files:
        if el.lower().endswith('.jpeg') or el.lower().endswith('.jpg'):
            result.append(el)
    return result

def getBlankResults():
    result = []
    for i in range(60):
        result.append('')
    return result

def drawQuestionsResults(image, questions):
    for el in questions: 
        x = int(el.coord.x)
        y = int(el.coord.y)
        text = 'NI'
        if el.num != 0 and el.ans != '':
            text = str(el.num) + ' ' + str(el.ans)
        border = 2
        font, font_scale, font_thickness = cv2.FONT_HERSHEY_COMPLEX, 0.7, 2
        
        text_size, _ = cv2.getTextSize(text, font, font_scale, font_thickness)
        text_w, text_h = text_size
        cv2.rectangle(image, (x-(text_w//2)-border,y+(text_h//2)+border), (x+(text_w//2)+border,y-(text_h//2)-border), (0,0,0), -1)
        cv2.putText(image, text, (x-(text_w//2),y+(text_h//2)), font, font_scale, (0,255,0), font_thickness)
    
    return image

def validatePosition(position, questionsCoordinates):
    rangeOfPixels = 20          # range of pixels for each row or column of questions
    minItensInRows = 10         # min of itens to consider that point valid
    minItensInColumns = 15      # min of itens to consider that point valid
    
    validX = False
    rangeX = list(filter(lambda q: abs(q.x-position.x) <= rangeOfPixels, questionsCoordinates))
    if len(rangeX)-minItensInColumns >= 0:
        validX = True
    
    validY = False
    rangeY = list(filter(lambda q: abs(q.y-position.y) <= rangeOfPixels, questionsCoordinates))
    if len(rangeY)-minItensInRows >= 0:
        validY = True

    return [validX and validY, rangeX, rangeY]

def validateQuestionsPosition(questionsCoordinates):
    result = []
    for el in questionsCoordinates:
        valid = validatePosition(el, questionsCoordinates)[0]
        if valid:
            result.append(el)
    return result

def getQuestionAndAnswerFromMap(coord):
    ans = ''
    question = 0
    section = 0
    for el in answersMap:
        if abs(el.x-coord.x) <= 15:
            ans = el.ans
            section = el.section
    
    for el in questionsMap:
        if abs(el.y-coord.y) <= 15 and el.section == section:
            question = el.number
    return [question, ans]

def updateMaps(rows, columns):   
    answersMap.append(AnsMap(columns[0].x, 'A', 0))
    answersMap.append(AnsMap(columns[1].x, 'B', 0))
    answersMap.append(AnsMap(columns[2].x, 'C', 0))
    answersMap.append(AnsMap(columns[3].x, 'D', 0))
    answersMap.append(AnsMap(columns[4].x, 'E', 0))
    answersMap.append(AnsMap(columns[5].x, 'A', 1))
    answersMap.append(AnsMap(columns[6].x, 'B', 1))
    answersMap.append(AnsMap(columns[7].x, 'C', 1))
    answersMap.append(AnsMap(columns[8].x, 'D', 1))
    answersMap.append(AnsMap(columns[9].x, 'E', 1))
    answersMap.append(AnsMap(columns[10].x, 'A', 2))
    answersMap.append(AnsMap(columns[11].x, 'B', 2))
    answersMap.append(AnsMap(columns[12].x, 'C', 2))
    answersMap.append(AnsMap(columns[13].x, 'D', 2))
    answersMap.append(AnsMap(columns[14].x, 'E', 2))
    
    for section in range(3):
        for i in range(len(rows)):
            questionsMap.append(QMap(rows[i].y, i+1+(section*20), section))
    
    return

def getColumnSection(answer):
    section = 0
    for i in range(len(answersMap)):
        if abs(answersMap[i].x-answer.x) <= 15:
            section = answersMap[i].section
    return section

def updateResults(answers, questions):
    results = getBlankResults()
    validAnswers = []
    validQuestions = validateQuestionsPosition(questions)
    
    for answer in answers:
        [valid, rangeX, rangeY] = validatePosition(answer, validQuestions)
        if valid:
            rows = sorted(rangeX, key=lambda r: r.y)
            columns = sorted(rangeY, key=lambda c: c.x)
            validAnswers.append(Question(answer))
            
            mapsNotFilled = (len(answersMap) == 0) and (len(questionsMap) == 0)
            if len(rows) == 20 and len(columns) == 15 and mapsNotFilled:
                updateMaps(rows, columns)
    
    # Process the results
    for valid_answer in validAnswers:
        [number, ans] = getQuestionAndAnswerFromMap(valid_answer.coord)
        if number > 0:
            if results[number - 1] != '':
                ans = 'ANULADA'
            valid_answer.num = number
            valid_answer.ans = ans
            results[number - 1] = ans
    
    return [results, validAnswers]

def getResults(image):
    image_gray, image_RGB = preprocess_image(image=image)
    detected_answers, detected_questions = detect_elements(image_gray=image_gray)
    
    if detected_answers is not None and detected_questions is not None: 
        answersCoordinates = []
        questionsCoordinates = []
        
        for el in detected_questions[0, :]:
            questionsCoordinates.append(Coordinates(el[0], el[1]))
        
        for el in detected_answers[0, :]:
            answersCoordinates.append(Coordinates(el[0], el[1]))
        print('Detected answered circles in', fileName, ':', len(answersCoordinates))
        
        result, validAns = updateResults(answersCoordinates, questionsCoordinates) 
        print('Valid answered circles in', fileName, ':', len(validAns), '\n')
    
        resultImage = drawQuestionsResults(image_RGB, validAns) 
        return [result, resultImage]

def getInfoUntilNextWordWithChar(char, line, initialIndex):
    line = line.upper()
    charIndex = line.find(char, initialIndex)
    
    if charIndex != -1:
        result = line[initialIndex:charIndex]
        return ' '.join(result.split(' ')[:-1])
    else:
        return line[initialIndex:]

def getTexts(image, resultImage):
    name, group = '', ''
    imageGray = cv2.cvtColor(image, cv2.COLOR_RGB2GRAY)[200:350]
        
    textLines = pytesseract.image_to_string(imageGray, config=r"--psm 3 osm 3").split('\n')
    for line in textLines:
        line = line.upper()
        
        index = line.find('NOME:')
        if index != -1:
            name = getInfoUntilNextWordWithChar(':', line.replace('NOME:', ''), index)
            name = name.strip()
        
        index = line.upper().find('TURMA:')
        if index != -1:
            group = getInfoUntilNextWordWithChar(':', line.replace('TURMA:', ''), index)
            group = group.strip()
    
    font, font_scale, font_thickness = cv2.FONT_HERSHEY_COMPLEX, 0.7, 2
    cv2.putText(resultImage, 'Nome: ' + str(name), (10,30), font, font_scale, (0,255,0), font_thickness)
    cv2.putText(resultImage, 'Turma: ' + str(group), (10,50), font, font_scale, (0,255,0), font_thickness)
    return [name, group, resultImage]

def writeStudentResults(cells, fileName, name, group, results, cell):
    cells.append(Cell(cell.row, cell.col, fileName))
    cells.append(Cell(cell.row, cell.col+1, len(list(filter(lambda r: r != '', results)))))
    cells.append(Cell(cell.row, cell.col+2, name))
    cells.append(Cell(cell.row, cell.col+3, group))
    
    for i in range(len(results)):    
        cells.append(Cell(cell.row, cell.col+4+i, results[i]))
    return

def writeHeader(cells, cell):
    cells.append(Cell(cell.row, cell.col, 'Nome do arquivo'))
    cells.append(Cell(cell.row, cell.col+1, 'Respostas registradas'))
    cells.append(Cell(cell.row, cell.col+2, 'Nome'))
    cells.append(Cell(cell.row, cell.col+3, 'Turma'))
    
    for i in range(60):    
        cells.append(Cell(cell.row, cell.col+4+i, 'Questão ' + str(i+1)))
    return

if (__name__ == '__main__'):
    initialTime = time.process_time()
    sheet = client.open(documentName)
    cells = []
    currentPath = os.path.dirname(__file__)
    files = getJPEGItensInPath(currentPath)
    
    writeHeader(cells, Cell(1,1))    
    for i in range(len(files)):
        fileName = files[i]
        img = cv2.imread(os.path.join(currentPath, fileName))
        resizedImg = cv2.resize(img, (1012, 1280), interpolation = cv2.INTER_NEAREST)
    
        try:
            [results, resultImage] = getResults(resizedImg)
            [name, group, resultImage] = getTexts(img, resultImage)
            
            resultsPath = os.path.join(currentPath, 'results')
            if not os.path.exists(resultsPath):
                os.makedirs(resultsPath)
            newPath = os.path.join(resultsPath, fileName)
            ski.io.imsave(newPath, resultImage)
            
            writeStudentResults(cells, fileName, name, group, results, Cell(2+i,1))
        
        except TypeError:
            print('Erro na leitura da imagem')
            writeStudentResults(cells, fileName, 'Erro', 'Erro', getBlankResults(), Cell(2+i,1))
    
    # if len(cells) > 0:
    #     worksheet = sheet.get_worksheet(0)
    #     worksheet.update_cells(cells)
        
    print('Done in', time.process_time()-initialTime, 'seconds!', '------------------------------\n')