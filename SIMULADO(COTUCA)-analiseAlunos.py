import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread.cell import Cell

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('cursinhox-47bc2e88124b.json', scope)
client = gspread.authorize(credentials)
    
# get the instance of the Spreadsheet
documentName = 'Correção Simulado COTUCA - 26/05 (respostas)'
divs = ['Malala', 'Marie Curie', 'Externo']

sheet = client.open(documentName)
data = sheet.get_worksheet(0).get_all_values()[1:]

gabarito = []
ranks = []
for el in divs:
    ranks.append([])
for el in data:
    if el[3].upper() == 'GABARITO':
        gabarito = el[4:]
    for i in range(len(divs)):
        if el[3].upper() == divs[i].upper():
            ranks[i].append(el)


rankingSheet = sheet.worksheet("Classificação atualizada")
rankingValues = rankingSheet.get_all_values()[1:]
rankingCols = []
for i in range(len(divs)):
    rankingCols.append(rankingSheet.find("Classificação " + divs[i]).col)
rankingCols.append(rankingSheet.find("Classificação Geral").col)

def getRankStats(name):  
    classGeral = 0
    classGeralTotal = 0
    classTurma = 0
    classTurmaTotal = 0
    notasIguais = -1
    qdeAcertos = 0
    pointsCTC = 0
    for el in rankingValues:
        for d in range(len(rankingCols)-1):
            if el[rankingCols[d]+1] == name:
                classTurma = el[rankingCols[d]-1]
                j = rankingCols[d]-1
        if el[rankingCols[-1]+1] == name:
            classGeral = el[rankingCols[-1]-1]
            qdeAcertos = el[rankingCols[-1]]
            pointsCTC = el[rankingCols[-1]+2].split(' ')[0]

    for el in rankingValues:
        cturma = el[j].replace('º', '')
        if cturma != '' and int(cturma) > int(classTurmaTotal):
            classTurmaTotal = cturma
            
        cGeral = el[rankingCols[-1]-1].replace('º', '')
        if cGeral != '' and int(cGeral) > int(classGeralTotal):
            classGeralTotal = cGeral
        
        if el[rankingCols[-1]] == qdeAcertos:
            notasIguais += 1
        
    return [classGeral, classGeralTotal + 'º', classTurma, classTurmaTotal + 'º', notasIguais, qdeAcertos, pointsCTC]
        
def getPointsForSubject(answer):
    port = 0
    mat = 0
    quim = 0
    fis = 0
    bio = 0        
    for i in range(len(gabarito)):
        if gabarito[i].upper() == answer[i].upper():
            if i+1 <= 12:
                port += 1
            elif i+1 <= 24:
                mat += 1
            elif i+1 <= 28:
                quim += 1
            elif i+1 <= 32:
                fis += 1
            elif i+1 <= 36:
                bio += 1
    return [port, mat, quim, fis, bio]

def writeAnaliseAluno(cells, name, rankStats, pointsForSubject, cell):
    cells.append(Cell(cell.row, cell.col, 'Nome:'))
    cells.append(Cell(cell.row, cell.col+1, name))
    
    cells.append(Cell(cell.row+2, cell.col, 'Classificação geral'))
    cells.append(Cell(cell.row+2, cell.col+1, str(rankStats[0]) + ' / ' + str(rankStats[1])))
    
    cells.append(Cell(cell.row+3, cell.col, 'Classificação na turma'))
    cells.append(Cell(cell.row+3, cell.col+1, str(rankStats[2]) + ' / ' + str(rankStats[3])))

    cells.append(Cell(cell.row+4, cell.col, 'Alunos que acertaram a mesma quantidade de questões'))
    cells.append(Cell(cell.row+4, cell.col+1, rankStats[4]))

    cells.append(Cell(cell.row+5, cell.col, 'Quantidade total de acertos'))
    cells.append(Cell(cell.row+5, cell.col+1, rankStats[5]))
    
    cells.append(Cell(cell.row+6, cell.col, 'Pontuação COTUCA'))
    cells.append(Cell(cell.row+6, cell.col+1, str(rankStats[6]) + ' / 108'))

    subjects = ['Português', 'Matemática', 'Quimica', 'Física', 'Biologia']
    qdeQuestoes = [12, 12, 4, 4, 4] 
    cells.append(Cell(cell.row+8, cell.col, 'Matéria'))
    cells.append(Cell(cell.row+8, cell.col+1, 'Questões'))
    cells.append(Cell(cell.row+8, cell.col+2, 'Acertos'))
    for i in range(len(pointsForSubject)):
        cells.append(Cell(cell.row+9+i, cell.col, subjects[i]))
        cells.append(Cell(cell.row+9+i, cell.col+1, qdeQuestoes[i]))
        cells.append(Cell(cell.row+9+i, cell.col+2, pointsForSubject[i]))    

def newWorksheet(worksheetName):
    for el in sheet.worksheets():
        if el.title.upper() == worksheetName.upper():
            el.clear()
            return el
    return sheet.add_worksheet(title=worksheetName, rows=100, cols=100)

for i in range(len(divs)):
    cells = []
    for j in range(len(ranks[i])):
        name = ranks[i][j][2]
        rankStats = getRankStats(name)
        pointsForSubject = getPointsForSubject(ranks[i][j][4:])
        writeAnaliseAluno(cells, name, rankStats, pointsForSubject, Cell(1+(j*20),1))
    if len(cells) > 0:
        worksheet = newWorksheet("Analise - " + divs[i])
        worksheet.update_cells(cells)