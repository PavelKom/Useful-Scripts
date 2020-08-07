"""
Report Reader for Intermac Vertmax
"""

from tkinter import *
import glob
import os
import datetime
import tkcalendar  # https://github.com/j4321/tkcalendar
import random
from tkintertable import TableCanvas, TableModel  # https://github.com/dmnfarrell/tkintertable

reportFolder = ''
defaultCfg = 'C:/WNC/user/REPORTS'
if not os.path.exists('reportreader.cfg'):
    with open('reportreader.cfg', 'w+', encoding="utf-8") as cfg:
        cfg.write(defaultCfg)
        cfg.close()
        reportFolder = defaultCfg
else:
    with open('reportreader.cfg', 'r', encoding="utf-8") as cfg:
        reportFolder = cfg.read()
        cfg.close()


def toFixed(numObj, digits=0):
    return f"{numObj:.{digits}f}"


def getRandomColor():
    s = ""
    for i in range(3):
        h = random.randint(31, 223)
        s = (hex(h)).replace("0x", "") + s
        if h < 16:
            s = '0' + s
    s = "#" + s
    return s


monthToInt = {
    "Jan": 1,
    "Feb": 2,
    "Mar": 3,
    "Apr": 4,
    "May": 5,
    "Jun": 6,
    "Jul": 7,
    "Aug": 8,
    "Sep": 9,
    "Oct": 10,
    "Nov": 11,
    "Dec": 12}
intToMonth = {v: k for k, v in monthToInt.items()}


class PieceData(object):
    def __init__(self, ID=0, name=None, start=None, end=None, time=None, thickness=None, pieces=None, produced=None,
                 remained=None, state=None, date=None):
        self.id = ID
        self.programName = name
        self.start = start
        self.end = end
        self.time = time
        self.thickness = thickness
        self.pieces = pieces
        self.produced = produced
        self.remained = remained
        self.state = state
        self.date = date
        self.isFirstPiece = False
        self.isLastPiece = False
        self.startTime = datetime.datetime(
            year=2000, month=1, day=1, hour=0, minute=0, second=0)
        self.endTime = datetime.datetime.now()
        self.timeTime = self.endTime - self.startTime
        self.timeCalc()
        if self.programName is None:
            self.programName = '[ERROR]'

    def timeCalc(self):
        if self.start is None:
            self.startTime = self.startTime.replace(
                year=self.date.year, month=self.date.month, day=self.date.day, hour=0, minute=0, second=0)
        else:
            self.startTime = self.calculateTime(self.start)
        if self.end is None:
            self.endTime = self.endTime.replace(
                year=self.date.year, month=self.date.month, day=self.date.day, hour=23, minute=59, second=59)
        else:
            self.endTime = self.calculateTime(self.end)
        if self.time is None:
            self.timeTime = self.endTime - self.startTime
        else:
            tmpDelta = [0, 0, 0, 0]
            tmpStr1 = str(self.time).replace("d", " ").replace("h", " ").replace("\'", " ").replace("\"", "")
            tmpStr2 = tmpStr1.split(' ')
            for i in range(min(len(tmpStr2), 4)):
                tmpDelta[3 - i] = int(round(float(tmpStr2[-i - 1])))
            self.timeTime = datetime.timedelta(
                days=tmpDelta[0], hours=tmpDelta[1], minutes=tmpDelta[2], seconds=tmpDelta[3])

    def calculateTime(self, timeStr):
        words = timeStr.split(' ')
        tmpTime = [0, 0, 0, 1, 1, 2000]
        tmpTime[-1] = int(words[-1])
        tmpTime[-2] = monthToInt[words[1]]
        tmpTime[-3] = int(words[2])

        tTime = words[3].split(":")
        tmpTime[0] = int(tTime[2])
        tmpTime[1] = int(tTime[1])
        tmpTime[2] = int(tTime[0])

        timeDatetime = datetime.datetime(
            year=tmpTime[5], month=tmpTime[4], day=tmpTime[3],
            hour=tmpTime[2], minute=tmpTime[1], second=tmpTime[0]
        )
        return timeDatetime


class RRInterface(Frame):
    today = datetime.datetime.now().date()
    pieces = []

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.configure(width=1250, height=500)

        self.frCal = Frame(self)
        self.frCal.place(relx=0, rely=0)
        self.frCal.configure(bd=1, relief=RAISED)
        self.calWidget = tkcalendar.Calendar(
            self.frCal, showweeknumbers=False, locale="ru_RU",
            maxdate=self.today)
        self.calWidget.pack()
        self.calWidget.bind('<<CalendarSelected>>', self.getDate)

        self.dayDataFrame = Frame(self)
        self.dayDataFrame.grid_propagate(0)
        self.dayDataFrame.place(relx=0, rely=1, anchor=SW)
        self.dayDataFrame.configure(width=1250, height=300, bd=1, relief=RAISED)

        self.tableModel = TableModel()
        self.table = TableCanvas(self.dayDataFrame, cellwidth=300, model=self.tableModel, rowheight=25)
        self.table.show()

        self.drawFrame = Frame(self)
        self.drawFrame.grid_propagate(0)
        self.drawFrame.place(relx=1, rely=0, anchor=NE)
        self.drawFrame.configure(width=966, height=200, bd=1, relief=RAISED)

        self.createCanvas()

        self.dateList = []
        self.hourUsed = [0 for i in range(24)]

        self.strInfo = StringVar()
        self.labelInfo = Label(self, textvariable=self.strInfo, width=30, height=1, bg='white', bd=1, relief=RAISED,
                               font='Arial 10')
        self.strInfo.set('Test')
        self.labelInfo.place(x=0, y=175)

        self.createFileList()
        self.createTable()
        self.readReportFile(self.today)

    def createCanvas(self):
        self.drawCanvas = Canvas(self.drawFrame, bg='white')
        self.drawCanvas.place(x=0, y=0)
        self.drawCanvas.configure(width=960, height=194)
        for i in range(24):
            self.drawCanvas.create_line((i + 1) * 40, 0, (i + 1) * 40, 180, fill='black')
            self.drawCanvas.create_text((i + 1) * 40 - 20, 90, text=str(i), font="Verdana 14")
        self.drawCanvas.create_line(0, 180, 960, 180, fill='black')

    def getDate(self, event=None):
        date = self.calWidget.selection_get()
        self.readReportFile(date)

    def createFileList(self):
        fileList = glob.glob(reportFolder+"\\*.REP")
        for file in fileList:
            file = file.replace(reportFolder+"\\", "")
            if int(file[1:5]) < 2000:
                continue
            self.dateList.append(
                datetime.date(year=int(file[1:5]), month=monthToInt.get(file[5:8]), day=int(file[8:10])))
        self.dateList.sort()
        if len(self.dateList) > 0:
            self.calWidget.configure(mindate=self.dateList[0])
        else:
            self.calWidget.configure(mindate=datetime.datetime.now())

    def createTable(self):
        data = {
            'def': {
                'Program name': 'test',
                'Start': '00:00:00',
                'End': '23:59:59',
                'Time': '23:59:59',
                'Thickness': 99,
                'Pieces': 999,
                'Produced': 111,
                'Remained': 888,
                'State': 'Ok'}}
        self.tableModel.importDict(data)
        self.table.resizeColumn(1, 150)
        self.table.resizeColumn(2, 150)
        self.table.resizeColumn(3, 70)
        self.table.resizeColumn(4, 90)
        self.table.resizeColumn(5, 60)
        self.table.resizeColumn(6, 80)
        self.table.resizeColumn(7, 80)

    def readReportFile(self, date):
        # R1970Jan01.REP
        fileName = 'R' + str(date.year) + intToMonth[date.month]
        if date.day < 10:
            fileName = fileName + '0' + str(date.day)
        else:
            fileName = fileName + str(date.day)
        fileName = reportFolder + '\\' + fileName + '.REP'

        self.pieces.clear()
        self.drawCanvas.delete('piece')
        self.tableModel.deleteRows(range(self.tableModel.getRowCount()))
        self.table.redraw()
        if not os.path.exists(fileName):
            self.strInfo.set('File not exist')
            return

        with open(fileName, 'r', encoding="utf-8") as scannedFile:
            fileData = scannedFile.read()
            pieces = fileData.split('\n\n')
            for i in range(len(pieces)):
                lines = pieces[i].split('\n')
                lineDict = {}
                for line in lines:
                    ll = line.replace("\n", "").split('=')
                    if len(ll) < 2:
                        continue
                    lineDict[ll[0]] = ll[1]
                self.pieces.append(
                    PieceData(ID=len(self.pieces),
                              name=lineDict.get('PROGRAM NAME'),
                              start=lineDict.get('START'),
                              end=lineDict.get('END'),
                              time=lineDict.get('MACHINING TIME'),
                              thickness=lineDict.get('WORKPIECE THICKNESS'),
                              pieces=lineDict.get('NUMBER OF PIECES'),
                              produced=lineDict.get('NR OF PIECES PRODUCED'),
                              remained=lineDict.get('NR OF PIECES REMAINING'),
                              state=lineDict.get('STATE'),
                              date=date))
        counter = 0
        while counter < len(self.pieces):
            if self.pieces[counter].programName == '[ERROR]' and self.pieces[counter].id > 0:
                self.pieces.pop(counter)
            else:
                counter += 1
        self.strInfo.set(str(date) + "\t Pieces: " + str(len(self.pieces)))
        self.addDayDataToCanvas()
        self.addDayDataToTable()

    def addDayDataToCanvas(self):
        for i in range(len(self.hourUsed)):
            self.hourUsed[i] = 0

        for piece in self.pieces:
            if piece.programName == '[ERROR]' and piece.id > 0:
                continue
            start = piece.startTime
            end = piece.endTime
            pieceColor = getRandomColor()
            sX = start.hour * 40
            sY = 180 - (start.minute + start.second / 60) * 3
            eX = sX + 40
            eY = 180 - (end.minute + end.second / 60) * 3
            if end.hour != start.hour:
                self.hourUsed[end.hour] += \
                    (end - end.replace(hour=end.hour, minute=0, second=0)).seconds
                currHour = start.hour + 1
                finalHour = end.hour
                if finalHour < currHour:
                    finalHour = 23
                for hH in range(currHour, finalHour):
                    self.drawCanvas.create_rectangle(
                        hH * 40, 180, (hH + 1) * 40, 0, fill=pieceColor, tag="piece")
                    self.hourUsed[hH] += 3600
                tsX = finalHour * 40
                tsY = 180
                teX = tsX + 40
                teY = eY
                eY = 0
                self.drawCanvas.create_rectangle(tsX, tsY, teX, teY, fill=pieceColor, tag="piece")
                if start.hour < 23:
                    self.hourUsed[start.hour] += (
                            start.replace(hour=start.hour + 1, minute=0, second=0) - start).seconds
                else:
                    nextDay = start + datetime.timedelta(seconds=3600)
                    self.hourUsed[start.hour] += (nextDay.replace(hour=0, minute=0, second=0) - start).seconds
            else:
                self.hourUsed[start.hour] += (end - start).seconds

            self.drawCanvas.create_rectangle(sX, sY, eX, eY, fill=pieceColor, tag="piece")
        for i in range(24):
            used = toFixed(self.hourUsed[i] / 36, 2)
            tt = self.drawCanvas.create_text(i * 40 + 20, 188, text=used, tag="piece")
            if tt > 100000:
                self.drawCanvas.delete('all')
                self.drawCanvas.destroy()
                self.createCanvas()
                self.addDayDataToCanvas()

    def addDayDataToTable(self):
        for piece in self.pieces:
            data = {piece.id: {
                'Program name': piece.programName,
                'Start': str(piece.startTime),
                'End': str(piece.endTime),
                'Time': str(piece.timeTime),
                'Thickness': piece.thickness,
                'Pieces': piece.pieces,
                'Produced': piece.produced,
                'Remained': piece.remained,
                'State': piece.state}}
            self.tableModel.importDict(data)
        self.table.redraw()


root = Tk()
root.geometry("1250x500")
root.resizable(0, 0)
root.title("Report Reader")
interface = RRInterface(root)

root.mainloop()
