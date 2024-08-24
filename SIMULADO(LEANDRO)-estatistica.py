import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread.cell import Cell

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('cursinhox-47bc2e88124b.json', scope)
client = gspread.authorize(credentials)
    
## VARIÁVEIS
documentName = 'Correção Simulado LF - 30/06 (respostas)'
divs = ['Malala', 'Marie Curie']


## CLASSES
class AlunoInfo:
    def __init__(self, name, className, answers):
        self.name = name
        self.className = className
        self.answers = answers
        self.points = 0
        self.classGeral = '0'
        self.classGeralTotal = '0'
        self.classTurma = '0'
        self.classTurmaTotal = '0'
        self.notasIguais = -1
        self.notaFinal = 0
    def __repr__(self):
        return repr((self.name, self.className, self.answers))
    
class TestSubject:
    def __init__(self, name, numOfQuestions, limit):
        self.name = name
        self.numOfQuestions = numOfQuestions
        self.limit = limit
    def __repr__(self):
        return repr((self.name, self.numOfQuestions, self.limit))


## PROGRAMA
subjects = [
    TestSubject('Português', 10, 10),
    TestSubject('História', 10, 20),
    TestSubject('Geografia', 10, 30),
    TestSubject('Química', 7, 37),
    TestSubject('Física', 7, 44),
    TestSubject('Biologia', 6, 50),
    TestSubject('Matemática', 10, 60)
]

def setRankStats(aluno): 
    for el in rankingValues:
        for d in range(len(rankingCols)-1):
            if el[rankingCols[d]+1] == aluno.name:
                aluno.classTurma = el[rankingCols[d]-1]
                j = rankingCols[d]-1
        if el[rankingCols[-1]+1] == aluno.name:
            aluno.classGeral = el[rankingCols[-1]-1]
            aluno.notaFinal = el[rankingCols[-1]]

    for el in rankingValues:
        cturma = el[j]
        if cturma != '' and int(cturma.replace('º', '')) > int(aluno.classTurmaTotal.replace('º', '')):
            aluno.classTurmaTotal = cturma
            
        cGeral = el[rankingCols[-1]-1]
        if cGeral != '' and int(cGeral.replace('º', '')) > int(aluno.classGeralTotal.replace('º', '')):
            aluno.classGeralTotal = cGeral
        
        if el[rankingCols[-1]] == aluno.notaFinal:
            aluno.notasIguais += 1
    return   

def getPointsForSubject(answer):
    pointsArr = []
    for i in range(len(subjects)):
        pointsArr.append(0)    
    for i in range(len(gabarito)):
        if gabarito[i].upper() == answer[i].upper():
            for j in range(len(subjects)):
                if i+1 <= subjects[j].limit:
                    pointsArr[j] += 1
                    break
    return pointsArr

def writeAnaliseAluno(cells, aluno, cell):
    cells.append(Cell(cell.row, cell.col, 'Nome:'))
    cells.append(Cell(cell.row, cell.col+1, aluno.name))
    
    cells.append(Cell(cell.row+2, cell.col, 'Classificação geral'))
    cells.append(Cell(cell.row+2, cell.col+1, str(aluno.classGeral) + ' / ' + str(aluno.classGeralTotal)))
    
    cells.append(Cell(cell.row+3, cell.col, 'Classificação na turma'))
    cells.append(Cell(cell.row+3, cell.col+1, str(aluno.classTurma) + ' / ' + str(aluno.classTurmaTotal)))

    cells.append(Cell(cell.row+4, cell.col, 'Notas iguais'))
    cells.append(Cell(cell.row+4, cell.col+1, aluno.notasIguais))

    cells.append(Cell(cell.row+5, cell.col, 'Nota final'))
    cells.append(Cell(cell.row+5, cell.col+1, aluno.notaFinal))

    cells.append(Cell(cell.row+7, cell.col, 'Matéria'))
    cells.append(Cell(cell.row+7, cell.col+1, 'Questões'))
    cells.append(Cell(cell.row+7, cell.col+2, 'Acertos'))
    pointsForSubject = getPointsForSubject(aluno.answers)
    for i in range(len(pointsForSubject)):
        cells.append(Cell(cell.row+8+i, cell.col, subjects[i].name))
        cells.append(Cell(cell.row+8+i, cell.col+1, subjects[i].numOfQuestions))
        cells.append(Cell(cell.row+8+i, cell.col+2, pointsForSubject[i]))    

def newWorksheet(worksheetName):
    for el in sheet.worksheets():
        if el.title.upper() == worksheetName.upper():
            el.clear()
            return el
    return sheet.add_worksheet(title=worksheetName, rows=100, cols=100)

def writeAverages(data, columns):
    geralRank = []
    divsRank = []
    for el in data:
        if el[columns[-1]] != '':
            geralRank.append(int(el.replace('/60', '')))
        for c in columns:
            if 
    return

if (__name__ == '__main__'):
    sheet = client.open(documentName)
    answersData = sheet.get_worksheet(0).get_all_values()[1:]

    gabarito = []
    ranks = []
    for el in divs:
        ranks.append([])
    for el in answersData:
        if el[3].upper() == 'GABARITO':
            gabarito = el[4:]
        for i in range(len(divs)):
            if el[3].upper() == divs[i].upper():
                ranks[i].append(AlunoInfo(el[2], el[3], el[4:]))
    
    rankingSheet = sheet.worksheet("Classificação atualizada")
    rankingValues = rankingSheet.get_all_values()[1:]
    rankingCols = []
    for i in range(len(divs)):
        rankingCols.append(rankingSheet.find("Classificação " + divs[i]).col)
    rankingCols.append(rankingSheet.find("Classificação Geral").col)
    print(rankingCols)
    
    writeAverages(rankingValues, rankingCols)
    for i in range(len(divs)):
        cells = []
        for j in range(len(ranks[i])):
            aluno = ranks[i][j]
            setRankStats(aluno)
            writeAnaliseAluno(cells, aluno, Cell(1+(j*20),1))
        if len(cells) > 0:
            worksheet = newWorksheet("Analise - " + divs[i])
            worksheet.update_cells(cells)