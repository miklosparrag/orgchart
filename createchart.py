__author__ = 'miklos'

from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4,A1, landscape
from reportlab.platypus import Paragraph, Frame
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.colors import HexColor

import gdata.docs
import gdata.docs.service
import gdata.spreadsheet.service
import getpass

# Connect to Google
gd_client = gdata.spreadsheet.service.SpreadsheetsService()
gd_client.email = 'miklos@we7.com'
gd_client.password = getpass.getpass("password:")
gd_client.source = 'we7.com'
gd_client.ProgrammaticLogin()

q = gdata.spreadsheet.service.DocumentQuery()
q['title'] = 'Orgchart'
q['title-exact'] = 'true'
feed = gd_client.GetSpreadsheetsFeed(query=q)
spreadsheet_id = feed.entry[0].id.text.rsplit('/',1)[1]
feed = gd_client.GetWorksheetsFeed(spreadsheet_id)
worksheet_id1 = feed.entry[0].id.text.rsplit('/',1)[1]
worksheet_id2 = feed.entry[1].id.text.rsplit('/',1)[1]

rows = gd_client.GetListFeed(spreadsheet_id, worksheet_id1).entry
titlerows = gd_client.GetListFeed(spreadsheet_id, worksheet_id2).entry

styles = getSampleStyleSheet()
styleN = styles['Normal']
styleN.alignment = 1
styleN.fontSize = 6
styleN.leading = 8
styleS = getSampleStyleSheet()['Normal']
styleS.alignment = 0
styleS.fontSize = 8
styleS.leading = 8
styleH = getSampleStyleSheet()['Normal']
styleH.alignment = 1
styleH.fontSize = 7
styleH.leading = 8
styleH.fontName = "Helvetica-Bold"
styleB = getSampleStyleSheet()['Normal']
styleB.alignment = 2
styleB.fontSize = 6
styleB.leading = 7
styleT = getSampleStyleSheet()['Normal']
styleT.alignment = 1
styleT.fontSize = 7
styleT.fontName = "Helvetica-Bold"
styleT.leading = 7
styleTitle = getSampleStyleSheet()['Normal']
styleTitle.alignment = 1
styleTitle.fontSize = 15
styleTitle.fontName = "Helvetica-Bold"
styleTitle.leading = 18

BIG = False

boxw = 74
boxh = 38
spaceh = 80

if BIG:
    boxw = 50
    boxh = 50

colors = [
    [None, 'OPEN', 'CONTRACTOR', '#FFFFFF'],
    [None, None, 'CONTRACTOR', '#F6FA78'],
    ['OXFORD', 'OFFER', None, '#FFFFFF'],
    ['OXFORD', 'OPEN', None, '#FFFFFF'],
    ['OXFORD', None, None, '#94FA61'],
    [None, 'OPEN', None, '#FFFFFF'],
    [None, 'OFFER', None, '#FFFFFF'],
    [None, None, None, '#61F2FA'],
]

pageTitles = {}

for row in titlerows:
    pageTitles[row.custom['pageid'].text] = row.custom['title'].text

LondonFTEs = 0
OxfordFTEs = 0
Contractors = 0
Vacancies = 0

class Employee:
    def __init__(self, name='', title='', reportTo='', level=0, jobNo='', status='', startdate='', location='', team='', desk = '', visible = True):
        self.name = name
        self.title = title
        self.level = level
        self.reportTo = reportTo
        self.reports = []
        self.visible = visible
        self.jobNo = jobNo
        self.status = status
        self.startdate = startdate
        self.location = location
        self.team = team
        self.top = False
        self.line = False
        self.desk = desk

    def update(self, e):
        self.name = e.name
        self.title = e.title
        self.level = e.level
        self.reportTo = e.reportTo
        self.visible = e.visible
        self.jobNo = e.jobNo
        self.status = e.status
        self.startdate = e.startdate
        self.location = e.location
        self.team = e.team
        self.line = False
        self.desk = desk

def drawInfoBox(c, t,  x, y):
    lines = []
    c.setFillColor(HexColor('#E8E8E8'))
    c.setStrokeColor(HexColor('#E8E8E8'))
    c.rect(x - boxw / 2, y, boxw, boxh / 3.0, fill=1)
    f = Frame(x - boxw / 2, y, boxw, boxh / 3.0, leftPadding=1, bottomPadding=1, rightPadding=1, topPadding=1, showBoundary=1)
    lines.append(Paragraph(t, styleT))
    f.addFromList(lines,c)
    c.setStrokeColor(HexColor('#000000'))

def drawBox(c, e,  x, y):

    c.setFillColorRGB(1, 1, 1)

    for cl in colors:
        if cl[0] and e.location.upper() != cl[0]:
            continue
        if cl[1] and e.startdate.upper() != cl[1]:
            continue
        if cl[2] and e.status.upper() != cl[2]:
            continue
        c.setFillColor(HexColor(cl[3]))

        break

    c.setStrokeColor(HexColor('#E8E8E8'))
    c.rect(x - boxw / 2, y - boxh / 2, boxw, boxh, fill=1)
    f = Frame(x - boxw / 2, y - boxh / 2, boxw, boxh, leftPadding=1, bottomPadding=1, rightPadding=1, topPadding=3, showBoundary=1)
    f2 = Frame(x - boxw / 2, y - boxh + 21, boxw, 11, leftPadding=1, bottomPadding=1, rightPadding=2, topPadding=3, showBoundary=0)

    lines = []

    if e.startdate.upper() == 'OPEN':
        lines.append(Paragraph('OPEN-'+e.status, styleH))
    else:
        lines.append(Paragraph(e.name, styleH))

    lines.append(Paragraph(e.title, styleN))
    lines2 = []
    lines2.append(Paragraph(e.desk, styleB))
    f2.addFromList(lines2, c)
    f.addFromList(lines, c)
    c.setStrokeColor(HexColor('#000000'))

def drawSummaryBox(x, y):

    c.setFillColorRGB(1, 1, 1)

    c.setStrokeColor(HexColor('#E8E8E8'))
    c.rect(x, y-70, 70, 70, fill=1)
    f = Frame(x, y-70, 70, 70, leftPadding=5, bottomPadding=5, rightPadding=5, topPadding=5, showBoundary=1)

    lines = []
    lines.append(Paragraph("Summary:",styleH))
    lines.append(Paragraph('' , styleS))
    lines.append(Paragraph('London: %d' % LondonFTEs, styleS))
    lines.append(Paragraph('Oxford: %d' % OxfordFTEs, styleS))
    lines.append(Paragraph('Total: %d' % (LondonFTEs+OxfordFTEs), styleS))
    lines.append(Paragraph('Contractors: %d' % Contractors, styleS))
    lines.append(Paragraph('Total All: %d' % (LondonFTEs+OxfordFTEs+Contractors), styleS))
    lines.append(Paragraph('Vacancies: %d' % Vacancies, styleS))


    f.addFromList(lines, c)
    c.setStrokeColor(HexColor('#000000'))

def calcBoxEmpTree(e):
    if e.level == 1:
        return 1, len(e.reports) + 1
    if e.level == 0:
        return 1, 1
    h = 0
    w = 0
    for ee in e.reports:
        ww, hh = calcBoxEmpTree(ee)
        h += hh
        w += ww
    if len(e.reports) == 0:
        return 1,1
    return w, h

def drawEmployees(c, e, x, y, w, h):
    bw, bh = calcBoxEmpTree(e)
    LRLW = 7
    if BIG:
        LRLW = 3

    if e.visible and not e.top:
        c.line(x + w / 2.0, y - (spaceh-boxh) / 2.0, x + w / 2.0, y)

    if e.visible and e.level == 1:
        drawInfoBox(c, e.team, x + w / 2.0, y - spaceh / 2.4)
        y -= boxh / 3.0

    if e.visible and e.reports:
        if e.level>1:
            c.line(x + w / 2.0, y - boxh, x + w / 2.0, y - spaceh)
        else:
            c.line(x + w / 2.0 + boxw / 2.0, y - boxh, x + w / 2.0 + boxw / 2.0 + LRLW, y - boxh)
            c.line(x + w / 2.0 + boxw / 2.0 + LRLW, y - boxh, x + w / 2.0 + boxw / 2.0 + LRLW, y - boxh*2 )

    if e.visible:
        drawBox(c, e, x + w / 2.0, y - spaceh / 2.0)

    if e.line:
        c.line(x + w / 2.0, y - boxh / 2.0, x + w / 2.0, y - spaceh)
        c.line(x + w / 2.0, y - boxh / 2.0, x + w / 2.0, y)

    if e.level == 1:
        y -= boxh * 1.2
        for ee in e.reports:
            if ee.visible:
                drawBox(c, ee, x + w / 2.0, y - spaceh / 2.0)
                c.line(x + w / 2.0 + boxw / 2.0, y - boxh, x + w / 2.0 + boxw / 2.0 + LRLW, y - boxh)
                c.line(x + w / 2.0 + boxw / 2.0 + LRLW, y - boxh, x + w / 2.0 + boxw / 2.0 + LRLW, y )
                y -= boxh
    else:
        if bw:
            y -= spaceh
            uw = w / bw
            h -= spaceh
            fbw = 0
            first = True
            ox = x
            for ee in e.reports:
                bw, bh = calcBoxEmpTree(ee)
                if first:
                    fbw = bw
                    first = False
                drawEmployees(c, ee, x, y, uw * bw, h)
                x += uw * bw
            if e.visible and e.reports:
                c.line(ox + uw * (fbw / 2.0), y, x - (uw * bw) / 2.0, y)

def drawPage(c, page, title, info = False):

    width, height = landscape(A4)
    TOP = 35
    if BIG:
        width, height = landscape(A1)
        TOP = 100

    if info:
        drawSummaryBox(width-90, 100)

    employees = page.employees
    roots = [e for e in employees if e.level == page.maxlevel]

    mw = 0
    for e in roots:
        w, h = calcBoxEmpTree(e)
        mw += w

    x = 20
    uw = (width - 40) / mw
    for e in roots:
        w, h = calcBoxEmpTree(e)
        drawEmployees(c, e, x, height - TOP, uw * w, height - TOP)
        x += uw * w

    f = Frame(0, height-35, width, 20, leftPadding=1, bottomPadding=1, rightPadding=1, topPadding=1, showBoundary=0)
    lines = []
    lines.append(Paragraph(title, styleTitle))
    f.addFromList(lines, c)

    # stats
    lc = [0] * (page.maxlevel+1)


    for e in employees:
        if not e.visible:
            continue
        lc[e.level] += 1


    c.showPage()

class Page:

    def __init__(self):
        self.employees = []
        self.level1s = 0
        self.maxlevel = 0


    def addEmployee(self,ne):
        if ne.level == 1:
            self.level1s += 1
        self.maxlevel = max(self.maxlevel, ne.level)
        self.employees.append(ne)

    def calculate(self):
        for i in self.employees:
            self.processEmployee(i)

    def processEmployee(self, ne):
        try:
            r = next(e for e in self.employees if e.name == ne.reportTo and e.visible)
            while ne.level < r.level-1:
                if ne.level == 0:
                    ne.level += 1
                    continue
                x = Employee(level=ne.level + 1, visible=False)
                x.reportTo = 'x'
                x.line = True
                x.reports.append(ne)
                self.employees.append(x)
                ne = x
            r.reports.append(ne)
        except StopIteration:
            ne.top = True

pages = {}

for row in rows:

    if int(row.custom['org.chart'].text)== -1:
        continue

    status = row.custom['status'].text or ''
    location = row.custom['baselocation'].text or ''
    name = row.custom['name'].text or ''
    title = row.custom['title'].text or ''
    reportsto = row.custom['reportsto'].text or ''
    level = int(row.custom['org.chart'].text)
    startdate = row.custom['startdate'].text or ''
    team = row.custom['team'].text or ''
    desk = row.custom['desk'].text or ''
    inpage = row.custom['inpage'].text or ''

    if status.upper() == 'CONTRACTOR':
        Contractors += 1
    elif status.upper() == 'FTE':
        if location.upper() == 'LONDON':
            LondonFTEs += 1
        else:
            OxfordFTEs += 1
    else:
        Vacancies += 1

    if BIG:
        inpage = '1'
    for p in inpage.split(','):
        if not p:
            continue
        if not pages.has_key(p):
            pages[p] = Page()

        pages[p].addEmployee(Employee(name, title, reportsto, level, '', status, startdate, location, team, desk))

for p in pages.values():
    p.calculate()


if BIG:
    c = canvas.Canvas("/Users/miklos/Google Drive/Technology Org Charts/orgchartBIG.pdf", pagesize = landscape(A1))
else:
    c = canvas.Canvas("/Users/miklos/Google Drive/Technology Org Charts/orgchart.pdf", pagesize = landscape(A4))

k = pages.keys()
k.sort()
info = True
for i in k:
    drawPage(c, pages[i], pageTitles[i], info)
    info = False

c.save()
