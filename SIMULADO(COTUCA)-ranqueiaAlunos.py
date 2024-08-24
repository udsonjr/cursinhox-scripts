import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread.cell import Cell
from operator import attrgetter

scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name('cursinhox-47bc2e88124b.json', scope)
client = gspread.authorize(credentials)
    
# get the instance of the Spreadsheet
documentName = 'Correção Simulado COTUCA - 26/05 (respostas)'
divs = ['Malala', 'Marie Curie', 'Externo']

class AlunoInfo:
    def __init__(self, name, className, points, port, mat, cien, total):
        self.name = name
        self.className = className
        self.points = points
        self.port = port
        self.mat = mat
        self.cien = cien
        self.total = total
    def __repr__(self):
        return repr((self.name, self.className, self.points, self.port, self.mat, self.cien, self.total))

sheet = client.open(documentName)
data = sheet.get_worksheet(0).get_all_values()[1:]

gabarito = []
for el in data:
    if el[3].upper() == 'GABARITO':
        gabarito = el[4:]
        data.remove(el)
        
lenGabarito = 0
for el in gabarito:
    if el != '':
        lenGabarito += 1
        
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
    return [port, mat, quim + fis + bio]

ranks = []
for el in divs:
    ranks.append([])
geralRank = []

for el in data:
    [port, mat, cien] = getPointsForSubject(el[4:])
    total = 3*port + 4*mat + 2*cien
    points = port + mat + cien
    alunoInfo = AlunoInfo(el[2], el[3], points, port, mat, cien, total)
    for i in range(len(divs)):
        if alunoInfo.className.upper() == divs[i].upper():
            ranks[i].append(alunoInfo)
    geralRank.append(alunoInfo)
    
for i in range(len(divs)):
    ranks[i] = sorted(ranks[i], key=attrgetter('total', 'mat', 'port', 'cien'), reverse=True)
geralRank = sorted(geralRank, key=attrgetter('total', 'mat', 'port', 'cien'), reverse=True)

def insertPosition(sortedRank):
    lastPosition = 1
    newRank = []
    if len(sortedRank) > 0:
        newRank.append([lastPosition, sortedRank[0]])
        for i in range(1, len(sortedRank)):
            if sortedRank[i].total == sortedRank[i-1].total and sortedRank[i].port == sortedRank[i-1].port and sortedRank[i].mat == sortedRank[i-1].mat and sortedRank[i].cien == sortedRank[i-1].cien:
                newRank.append([lastPosition, sortedRank[i]])
            else:
                newRank.append([i+1, sortedRank[i]])
            lastPosition = newRank[-1][0]
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
        cells.append(Cell(row+i+1, col, str(rank[i][0]) + 'º'))
        cells.append(Cell(row+i+1, col+1, str(rank[i][1].points) + ' / ' + str(lenGabarito)))
        cells.append(Cell(row+i+1, col+2, rank[i][1].name))        
        cells.append(Cell(row+i+1, col+3, str(rank[i][1].total) + ' / ' + str(rank[i][1].port) + '-' + str(rank[i][1].mat) + '-' + str(rank[i][1].cien)))
    worksheet.update_cells(cells)

for i in range(len(divs)):
    writeRank(ranks[i], 1, (i*5)+1, "Classificação " + divs[i])
writeRank(geralRank, 1, (len(ranks)*5)+1, "Classificação Geral")