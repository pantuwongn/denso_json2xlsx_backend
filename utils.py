from PIL import Image, ImageDraw
from datetime import datetime
from openpyxl.styles.borders import Border, Side
from openpyxl.styles.alignment import Alignment
from openpyxl.styles.fonts import Font
from openpyxl.drawing.text import TextField

outputDir = 'output'
bottomRightBorder = Border(
    bottom = Side(style='thin'),
    right = Side(style='thin'),
)
rightBorder = Border(
    right = Side(style='thin'),
)
bottomBorder = Border(
    bottom = Side(style='thin'),
)

leftCenterAlignment = Alignment(
    horizontal='left',
    vertical='center',
    wrap_text=True
)
topCenterAlignment = Alignment(
    horizontal='center',
    vertical='top',
    wrap_text=True
)
centerCenterAlignment = Alignment(
    horizontal='center',
    vertical='center',
    wrap_text=True
)
topLeftAlignment = Alignment(
    horizontal='left',
    vertical='top',
    wrap_text=True
)
headerNormalStyle = Font(name='CordiaUPC', size=12, color='000000')
textNormalStyle = Font(name='CordiaUPC', size=10, color='000000')

def getOutputFilePath(fileName):
    return '{outputDir}/{fileName}.xlsx'.format(
        outputDir = outputDir,
        fileName = fileName
    )

def chunk(iterable, chunk_size):
    # Initialize an empty list to store the chunks
    chunks = []
    
    # Iterate over the iterable in chunks of the specified size
    for i in range(0, len(iterable), chunk_size):
        # Get the chunk
        chunk = iterable[i:i + chunk_size]
        # Add the chunk to the list of chunks
        chunks.append(chunk)
    
    return chunks

def drawVerticalDashedLine(height: int):
    img = Image.new("RGB", (1, height), (255, 255, 255))

    d = ImageDraw.Draw(img)
    cur_y = 0
    space = 4
    length = 4
    for y in range(cur_y, height, length + space):
        d.line([0, y, 0, y + length], fill=(0, 0, 0), width=1)
    # img.show()
    filename = "temp/" + datetime.now().strftime("%Y%m%d%H%M%S%f") + ".png"
    # print(filename)
    img.save(filename)
    return filename
    
