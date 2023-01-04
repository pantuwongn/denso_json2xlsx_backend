from openpyxl import load_workbook
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.drawing.image import Image
from openpyxl.drawing.text import TextField
from openpyxl.utils.units import cm_to_EMU, pixels_to_EMU, EMU_to_pixels
from utils import (
    getOutputFilePath,
    drawVerticalDashedLine,
    chunk,
    leftCenterAlignment,
    centerCenterAlignment,
    topCenterAlignment,
    topLeftAlignment,
    textNormalStyle,
    bottomBorder,
    bottomRightBorder
)

itemChunkSize = 17

timingConnectorPath = 'images/timing/check-process.png'

counterPathMap = {
    1: 'images/counter/1.png',
    2: 'images/counter/2.png',
    3: 'images/counter/3.png',
    4: 'images/counter/4.png',
    5: 'images/counter/5.png',
    6: 'images/counter/6.png',
    7: 'images/counter/7.png',
    8: 'images/counter/8.png',
    9: 'images/counter/9.png',
    10: 'images/counter/10.png',
    11: 'images/counter/11.png',
    12: 'images/counter/12.png',
    13: 'images/counter/13.png',
    14: 'images/counter/14.png',
    15: 'images/counter/15.png',
    16: 'images/counter/16.png',
}

checkTimingSymbolPathMap = {
    'None': 'images/timing/check-no-record.png',
    'Check sheet': 'images/timing/check-record.png',
    'Record sheet': 'images/timing/check-record.png',
    'x-R chart': 'images/timing/check-control-chart.png',
    'xbar-R chart': 'images/timing/check-control-chart.png',
    'x-Rs chart': 'images/timing/check-control-chart.png',
}

scSymbolPathMap = {
    'C-none': 'images/symbols/C-none.png',
    'S-circle': 'images/symbols/S-circle.png',
    'S-diamond': 'images/symbols/S-diamond.png',
    'F-circle': 'images/symbols/F-circle.png',
    'F-triangle': 'images/symbols/F-triangle.png',
    'RW-rectangle': 'images/symbols/RW-rectangle.png',
    'SP-circle': 'images/symbols/SP-circle.png'
}

c2e = cm_to_EMU
p2e = pixels_to_EMU
cellh = lambda x: c2e(x * 0.48)
cellw = lambda x: c2e(x * 1.1)

def getSCSymbolList(scSymbolList: list, rowStart: int, maxRow: int):
    imgList = list()
    symbolTotal = len(scSymbolList)

    assert symbolTotal <= maxRow, 'SC Symbol out of bound'

    for i, scSymbol in enumerate(scSymbolList):
        symbolPath = scSymbolPathMap.get('{}-{}'.format(scSymbol['character'], scSymbol['shape']), None)
        if symbolPath is None:
            raise KeyError('Unregistered sc symbol, {}-{}'.format(scSymbol['character'], scSymbol['shape']))

        if symbolTotal == 1:
            symbolImg = drawImage(Image(symbolPath), rowStart, 2, 0, 10)
        else:
            symbolImg = drawImage(Image(symbolPath), rowStart + i -1, 2, 5*i, 10)
        imgList.append(symbolImg)
    return imgList

def getTotalSCSymbolList(itemList: list):
    scSymbolDict = dict()
    scSymbolCountDict = dict()
    for item in itemList:
        scSymbolList = item['sc_symbols']
        for scSymbol in scSymbolList:
            symbolHash = '{}-{}'.format(scSymbol['character'], scSymbol['shape'])
            scSymbolDict[symbolHash] = scSymbol

            if scSymbolCountDict.get(symbolHash, None) is None:
                scSymbolCountDict[symbolHash] = 1
            else:
                scSymbolCountDict[symbolHash] += 1

    totalSymbolList = list(scSymbolDict.values())

    imgList = list()
    for i, scSymbol in enumerate(totalSymbolList):
        symbolHash = '{}-{}'.format(scSymbol['character'], scSymbol['shape'])
        symbolPath = scSymbolPathMap.get(symbolHash, None)

        if symbolPath is None:
            raise KeyError('Unregistered sc symbol, {}-{}'.format(scSymbol['character'], scSymbol['shape']))

        symbolImg = drawImage(Image(symbolPath), 6, 12, 0, 35 * i)
        counterImg = drawImage(Image(counterPathMap[scSymbolCountDict[symbolHash]]), 6, 12, 12, (35 * i) + 23)
        imgList.append(symbolImg)
        imgList.append(counterImg)
    
    return imgList
    

def getProcessCapability(capabilityDict: dict):
    x_bar = 'xbar : {}'.format(capabilityDict['x_bar']) if capabilityDict['x_bar'].strip() != '' else ''
    cpk = 'cpk : {}'.format(capabilityDict['cpk']) if capabilityDict['cpk'].strip() != '' else ''

    result = ''
    if x_bar != '':
        result = '{}{}\n'.format(result, x_bar)
    if cpk != '':
        result = '{}{}'.format(result, cpk)

    return result

def getParameter(parameterDict: dict):
    limitType = parameterDict['limit_type']
    if limitType == 'None':
        return parameterDict['parameter']


    def _appendTextIfExist(targetStr: str, dataDict: dict):
        def doAppendTextIfExist(key: str):
            finalText = targetStr
            if dataDict.get(key, '').strip() != '':
                finalText = '{}{}'.format(targetStr, dataDict[key])
            return finalText
        return doAppendTextIfExist

    result = '{}\n'.format(parameterDict['parameter'])
    appendTextIfExist = _appendTextIfExist(result, parameterDict)
    result = appendTextIfExist('prefix')
    result = appendTextIfExist('main')
    result = appendTextIfExist('suffix')
    result = appendTextIfExist('tolerance_up')
    result = appendTextIfExist('tolerance_down')
    result = appendTextIfExist('unit')

    return result

def getInterval(controlMethodDict: dict):
    intervalText = controlMethodDict['interval']
    if controlMethodDict['100_method'] == 'Auto check':
        intervalText = '100%\n{}'.format(intervalText)

    if controlMethodDict['sample_no'] > 1:
        intervalText = '{}\nn=({})'.format(intervalText, controlMethodDict['sample_no'])
    return intervalText

def getControlMethodDetail(controlMethodDict: dict):
    if controlMethodDict.get('calibration_interval', '') != '':
        return 'Calibration'
    return ''

def getControlMethod(itemDict: dict):
    if (itemDict['control_method']['100_method'] == 'None'):
        return itemDict['control_item_type']
    return itemDict['control_method']['100_method']

def getMeasurement(itemDict: dict):
    finalText = itemDict['measurement']
    if itemDict['readability'] != '' or itemDict['parameter']['unit'] != '':
        finalText ='{} ({} {})'.format(
            finalText,
            itemDict['readability'],
            itemDict['parameter']['unit']
        )
    return finalText

def drawImage(img, row, col, rowOff, colOff):
    h, w = img.height, img.width
    size = XDRPositiveSize2D(p2e(w), p2e(h))
    marker = AnchorMarker(
        row=row,
        col=col,
        rowOff=p2e(rowOff),
        colOff=p2e(colOff)
    )
    img.anchor = OneCellAnchor(marker, size)
    return img

def getVerticalDashLine(height, row, col, rowOff, colOff):
    img = Image(drawVerticalDashedLine(EMU_to_pixels(c2e(height) * 0.45)))
    return drawImage(img, row, col, rowOff, colOff)

def getHorizontalDashLine(row, col, rowOff, colOff):
    return drawImage(
        Image('images/timing/dash-main-to-branch.png'),
        row,
        col,
        rowOff,
        colOff)

def getCheckProcess(row, col, rowOff, colOff):
    return drawImage(
        Image('images/timing/check-process.png'),
        row,
        col,
        rowOff,
        colOff)

def getCheckTimingSymbol(checkTiming, row, col, rowOff, colOff):
    symbolPath = checkTimingSymbolPathMap.get(checkTiming, None)
    if symbolPath is None:
        raise KeyError('Unregistered check timing type, {}'.format(checkTiming))
    img = Image(symbolPath)
    return drawImage(img, row, col, rowOff, colOff)

class PCSForm:
    def __init__(self, templatePath: str, dataDict: dict):
        self.templatePath = templatePath
        self.templateSheetName = 'empty'

        self.dataDict = dataDict
        self._initializeTemplateWorkbook()

    def _initializeTemplateWorkbook(self):
        self.workbook = load_workbook(filename = self.templatePath)

    def generate(self, fileName: str):
        headerDict = self.dataDict
        processList = self.dataDict['processes']

        templateSheet = self.workbook[self.templateSheetName]
        self._writeFormHeader(headerDict, templateSheet)

        processWithCheckTimingList = self._computeProcessCheckTiming(processList)

        totalProcess = len(processWithCheckTimingList)
        pageCount = 1
        for i, processDict in enumerate(processList):
            itemChunkList = chunk(processDict['items'], itemChunkSize)
            totalChunk = len(itemChunkList)-1 if len(itemChunkList) > 1 else 1
            for j, itemChunk in enumerate(itemChunkList):
                itemSheet = self.workbook.copy_worksheet(templateSheet)
                self._writeFormProcess(
                    pageCount, totalProcess + totalChunk,
                    j+1, totalChunk,
                    processDict,
                    itemSheet)
                itemSheet.title = 'process-{}-{}'.format(
                    i+1,
                    j+1
                )
                self._writeProcessItem(
                    itemChunkSize * (j),
                    itemSheet,
                    itemChunk
                )
                pageCount = pageCount + 1
        
        self._saveWorkbook(fileName)

    def _computeProcessCheckTiming(self, processList: list):
        processWithCheckTimingList = list()
        for i, processDict in enumerate(processList):
            processCheckTimingDict = dict(**processDict)
            processWithCheckTimingList.append(processCheckTimingDict)
        return processWithCheckTimingList

    def _writeFormHeader(self, headerDict: dict, sheet: Worksheet):
        #   Write check box
        sheet.cell(row=3, column=14).value = '❑    \t  Prototype'
        sheet.cell(row=4, column=14).value = '❑    \t  Pre-Launch'
        sheet.cell(row=5, column=14).value = '❑    \t  Production'

        sheet.cell(row=7, column=1).value = headerDict['line']
        sheet.cell(row=7, column=1).alignment = leftCenterAlignment
        sheet.cell(row=7, column=1).font = textNormalStyle
        sheet.cell(row=7, column=8).value = headerDict['assy_name']
        sheet.cell(row=7, column=8).alignment = leftCenterAlignment
        sheet.cell(row=7, column=8).font = textNormalStyle
        sheet.cell(row=9, column=8).value = headerDict['part_name']
        sheet.cell(row=9, column=8).alignment = leftCenterAlignment
        sheet.cell(row=9, column=8).font = textNormalStyle
        sheet.cell(row=9, column=13).value = headerDict['customer']
        sheet.cell(row=9, column=13).alignment = centerCenterAlignment
        sheet.cell(row=9, column=13).font = textNormalStyle

        sheet.cell(row=63, column=7).value = '                   Issue to ❑ Insp.    ❑ Prod.(___________)'

    def _writeFormProcess(self, idx: int, total: int, subIdx: int, subTotal: int, processDict: dict, sheet: Worksheet):
        #   Add denso logo
        densoIconImage = Image('images/denso-logo.png')
        h, w = densoIconImage.height, densoIconImage.width
        size = XDRPositiveSize2D(p2e(w), p2e(h))
        marker = AnchorMarker(
            row=0,
            col=0,
        )
        densoIconImage.anchor = OneCellAnchor(marker, size)
        sheet.add_image(densoIconImage)

        sheet.cell(row=2, column=15).value = 'Page  {} / {}'.format(
            idx,
            total
        )

        sheet.cell(row=9, column=1).value = '{}                  {} / {}'.format(
            processDict['name'],
            subIdx,
            subTotal
        )
        sheet.cell(row=9, column=1).alignment = leftCenterAlignment
        sheet.cell(row=9, column=1).font = textNormalStyle

    def _writeProcessItem(self, startNumber:int, sheet: Worksheet, itemList: list):
        startRow = 12
        rowStep = 3
        startSeparatorColumn = 3
        endSeparatorColumn = 15

        #   Dash line
        vertDashImg = getVerticalDashLine(len(itemList) * 3, 11, 1, -1, -17)
        sheet.add_image(vertDashImg)

        for i, item in enumerate(itemList):
            #   Cell merging
            sheet.merge_cells('E{}:H{}'.format(startRow + (rowStep * i), startRow + (rowStep * i)+1))
            sheet.merge_cells('E{}:H{}'.format(startRow + (rowStep * i) + 2, startRow + (rowStep * i) + 2))
            sheet.merge_cells('I{}:I{}'.format(startRow + (rowStep * i), startRow + (rowStep * i)+1))
            sheet.merge_cells('J{}:J{}'.format(startRow + (rowStep * i), startRow + (rowStep * i)+1))
            sheet.merge_cells('K{}:K{}'.format(startRow + (rowStep * i), startRow + (rowStep * i)+1))
            sheet.merge_cells('L{}:L{}'.format(startRow + (rowStep * i), startRow + (rowStep * i)+2))
            sheet.merge_cells('M{}:N{}'.format(startRow + (rowStep * i), startRow + (rowStep * i) + 2))
            sheet.merge_cells('O{}:O{}'.format(startRow + (rowStep * i), startRow + (rowStep * i) + 2))
            sheet.merge_cells('M7:O7')
            sheet.merge_cells('A9:G9')

            #   Cell bordering
            sheet.cell(row=startRow + (rowStep * i) + 1, column=5).border = bottomBorder
            sheet.cell(row=startRow + (rowStep * i) + 1, column=9).border = bottomBorder
            sheet.cell(row=startRow + (rowStep * i) + 1, column=10).border = bottomBorder
            sheet.cell(row=startRow + (rowStep * i) + 1, column=11).border = bottomBorder
            for j in range(endSeparatorColumn - startSeparatorColumn):
                sheet.cell(row=startRow + (rowStep * i) + 2, column=startSeparatorColumn + j + 1).border = bottomRightBorder

            #   Cell values
            sheet.cell(row=startRow + (rowStep * i), column=4).value = startNumber + (i + 1)
            sheet.cell(row=startRow + (rowStep * i), column=4).alignment = centerCenterAlignment
            sheet.cell(row=(startRow + (rowStep * i)), column=5).value = getParameter(item['parameter'])
            sheet.cell(row=(startRow + (rowStep * i) + 2), column=5).value = getMeasurement(item)
            sheet.cell(row=(startRow + (rowStep * i)), column=5).alignment = topLeftAlignment
            sheet.cell(row=(startRow + (rowStep * i)), column=9).value = getInterval(item['control_method'])
            sheet.cell(row=(startRow + (rowStep * i) + 2), column=9).value = item['control_method'].get('calibration_interval', '')
            sheet.cell(row=(startRow + (rowStep * i) + 2), column=9).alignment = centerCenterAlignment
            sheet.cell(row=(startRow + (rowStep * i)), column=9).alignment = topCenterAlignment
            sheet.cell(row=(startRow + (rowStep * i)), column=10).value = getControlMethod(item)
            sheet.cell(row=(startRow + (rowStep * i)), column=10).alignment = centerCenterAlignment
            sheet.cell(row=(startRow + (rowStep * i)) + 2, column=10).value = getControlMethodDetail(item['control_method'])
            sheet.cell(row=(startRow + (rowStep * i)) + 2, column=10).alignment = centerCenterAlignment
            sheet.cell(row=(startRow + (rowStep * i)), column=11).value = item['control_method']['in_charge']
            sheet.cell(row=(startRow + (rowStep * i)), column=11).alignment = centerCenterAlignment
            sheet.cell(row=(startRow + (rowStep * i)), column=12).value = getProcessCapability(item['initial_p_capability'])
            sheet.cell(row=(startRow + (rowStep * i)), column=13).value = item['remark']['remark']
            sheet.cell(row=(startRow + (rowStep * i)), column=13).alignment = topLeftAlignment
            sheet.cell(row=(startRow + (rowStep * i)), column=15).value = item['remark']['ws_no']
            sheet.cell(row=(startRow + (rowStep * i)), column=15).alignment = centerCenterAlignment

            vertDashImg = getVerticalDashLine(len(itemList) * 2, 13, 1, -1, -5)
            sheet.add_image(vertDashImg)

            #   Imaging
            scSymbolImgList = getSCSymbolList(item['sc_symbols'], startRow + (rowStep * i), 3)
            for scSymbol in scSymbolImgList:
                sheet.add_image(scSymbol)

            horizontalControlItemImg = getHorizontalDashLine(
                startRow + (rowStep * i), 1, 7.5, -4
            )
            sheet.add_image(horizontalControlItemImg)
            horizontalControlItemImg = getHorizontalDashLine(
                startRow + (rowStep * i), 1, 7.5, -10
            )
            sheet.add_image(horizontalControlItemImg)

            totalScSymbolList = getTotalSCSymbolList(itemList)
            for totalScSymbol in totalScSymbolList:
                sheet.add_image(totalScSymbol)

            controlItemSymbolImg = getCheckTimingSymbol(
                item['control_item_type'],
                startRow + (rowStep * i), 1, 0, 5)
            sheet.add_image(controlItemSymbolImg)

            checkProcessImage = getCheckProcess(
                startRow + (rowStep * i), 1, 0, -25
            )
            sheet.add_image(checkProcessImage)

    def _saveWorkbook(self, fileName: str):
        templateSheet = self.workbook[self.templateSheetName]
        self.workbook.remove(templateSheet)
        self.workbook.save(getOutputFilePath(fileName))
        
