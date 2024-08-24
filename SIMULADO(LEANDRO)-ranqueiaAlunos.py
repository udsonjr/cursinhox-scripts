import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread.cell import Cell
from operator import attrgetter

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
    def __repr__(self):
        return repr((self.name, self.className, self.answers, self.points))

class rankedAluno:
    def __init__(self, position, data):
        self.position = position
        self.data = data
    def __repr__(self):
        return repr((self.position, self.data))



## PROGRAMA
sheet = client.open(documentName)
data = sheet.get_worksheet(0).get_all_values()[1:]

gabarito = []
alunosGeral = []
for el in data:
    answers = el[4:]
    if el[3].upper() == 'GABARITO':
        gabarito = answers
    else:
        alunosGeral.append(AlunoInfo(el[2], el[3], el[4:]))
        
lenGabarito = 0
for el in gabarito:
    if el != '':
        lenGabarito += 1
        
def getPoints(answer, gabarito):
    points = 0
    for i in range(len(gabarito)):
        if gabarito[i].upper() == answer[i].upper():
            points += 1
    return points

ranks = []
for el in divs:
    ranks.append([])
geralRank = []
for aluno in alunosGeral:
    points = getPoints(aluno.answers, gabarito)
    aluno.points = points
    for i in range(len(divs)):
        if aluno.className.upper() == divs[i].upper():
            ranks[i].append(aluno)
    geralRank.append(aluno)
    
for i in range(len(divs)):
    ranks[i] = sorted(ranks[i], key=attrgetter('points'), reverse=True)
geralRank = sorted(geralRank, key=attrgetter('points'), reverse=True)

def insertPosition(sortedRank):
    lastPosition = 1
    newRank = []
    if len(sortedRank) > 0:
        newRank.append(rankedAluno(lastPosition, sortedRank[0]))
        for i in range(1, len(sortedRank)):
            if sortedRank[i].points == sortedRank[i-1].points:
                newRank.append(rankedAluno(lastPosition, sortedRank[i]))
            else:
                newRank.append(rankedAluno(i+1, sortedRank[i]))
            lastPosition = newRank[-1].position
    return newRank

for i in range(len(divs)):
    ranks[i] = insertPosition(ranks[i])
geralRank = insertPosition(geralRank)


#########################################################################

def newWorksheet(worksheetName):
    for el in sheet.worksheets():
        if el.title == worksheetName:
            el.clear()
            return el
    return sheet.add_worksheet(title=worksheetName, rows=100, cols=100)

worksheet = newWorksheet("Classificação atualizada")

def writeRank(rank, row, col, name):
    cells = []
    cells.append(Cell(row, col, name))
    for i in range(len(rank)):
        cells.append(Cell(row+i+1, col, str(rank[i].position) + 'º'))
        cells.append(Cell(row+i+1, col+1, str(rank[i].data.points) + ' / ' + str(lenGabarito)))
        cells.append(Cell(row+i+1, col+2, rank[i].data.name))
    worksheet.update_cells(cells)

for i in range(len(divs)):
    writeRank(ranks[i], 1, (i*4)+1, "Classificação " + divs[i])
writeRank(geralRank, 1, (len(ranks)*4)+1, "Classificação Geral")