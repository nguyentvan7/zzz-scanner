import os
import PIL.Image
import PIL.ImageColor
import PIL.ImageEnhance
import PIL.ImageGrab
import PIL.ImageOps
import win32gui
import PIL
import datetime
import pytesseract
import pyperclip

# x, y, w, h
# for 1440p
coords = {
    # 'discPopup': (793, 283, 1765, 1077),
    'setNameSlot': (836, 314, 1433, 380),
    'baseStatKey': (840, 602, 1076, 664),
    'baseStatValue': (1076, 602, 1276, 664),
    'subStat1Key': (840, 730, 1076, 797),
    'subStat1Value': (1076, 730, 1276, 797),
    'subStat2Key': (1300, 730, 1570, 797),
    'subStat2Value': (1570, 730, 1770, 797),
    'subStat3Key': (840, 810, 1087, 864),
    'subStat3Value': (1087, 810, 1287, 864),
    'subStat4Key': (1300, 810, 1570, 865),
    'subStat4Value': (1570, 810, 1770, 865),
    # 'rank': (1641, 498, 1691, 548)
}

def process(img):
    # Grayscale
    img = img.convert('L')
    # Threshold + invert
    img = img.point(lambda p: 0 if p > 200 else 255)
    # Mono
    img = img.convert('1')
    return img

def screenshot(bbox):
    return process(PIL.ImageGrab.grab(bbox=bbox, all_screens=True))

def saveImg(image, name):
    image.save('debug/{time}/{name}.png'.format(time = timestamp, name = name))

def saveScreenshot(name, bbox):
    img = screenshot(bbox)
    saveImg(img, name)
    return img

fullPath = r'./debug\full.png'
def loadAndProcess(bbox):
    return process(PIL.Image.open(fullPath).crop(bbox))

timestamp = datetime.datetime.now().strftime('%Y-%m-%dT%H%M%S')
debug = True

if __name__ == "__main__":
    while option := input("Press Enter to take screenshot. Input 'q' to quit.") != 'q':
        window = win32gui.FindWindow("UnityWndClass", "ZenlessZoneZero")
        if debug:
            path = os.path.join('debug', timestamp)
            if not(os.path.isdir(path)):
                os.makedirs(os.path.join('debug', timestamp))
        if window != 0 and fullPath != '':
            width, height = win32gui.GetClientRect(window)[2:]
            x, y = win32gui.ClientToScreen(window, (0, 0))

            if debug:
                saveScreenshot('full', (x, y, x + width, y + height))

        tesseractPath = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        pytesseract.pytesseract.tesseract_cmd = tesseractPath
        
        # TODO: Scale based on resolution/position of game
        stats = {}
        for name, bbox in coords.items():
            # Crop and process
            if fullPath != '':
                img = loadAndProcess(bbox)
            else:
                img = screenshot(bbox)
                img = process(screenshot)
            if debug:
                saveImg(img, name)

            # OCR
            ocrStr = pytesseract.image_to_string(img, lang='eng', config='--oem 1 --psm 6')
            cleanStr = ocrStr.strip().replace('\n', ' ')
            if name == "setNameSlot":
                bracket = cleanStr.find('[')
                setName = cleanStr[:bracket-1]
                stats['setName'] = setName
                slot = cleanStr[bracket+1: bracket+2]
                stats['slot'] = slot
            else:
                stats[name] = cleanStr
        match stats['baseStatKey']:
            case 'ATK' | 'DEF' | 'HP':
                if stats['baseStatValue'].endswith('%'):
                    stats['baseStatKey'] += ' %'
        match stats['subStat1Key']:
            case 'ATK' | 'DEF' | 'HP':
                if stats['subStat1Value'].endswith('%'):
                    stats['subStat1Key'] += ' %'
        match stats['subStat2Key']:
            case 'ATK' | 'DEF' | 'HP':
                if stats['subStat2Value'].endswith('%'):
                    stats['subStat2Key'] += ' %'
        match stats['subStat3Key']:
            case 'ATK' | 'DEF' | 'HP':
                if stats['subStat3Value'].endswith('%'):
                    stats['subStat3Key'] += ' %'
        match stats['subStat4Key']:
            case 'ATK' | 'DEF' | 'HP':
                if stats['subStat4Value'].endswith('%'):
                    stats['subStat4Key'] += ' %'
        print(stats)
        excelStr = '{}\t{}\t{}\t{}\t{}\t{}'.format(stats['setName'], stats['slot'], stats['baseStatKey'], stats['subStat1Key'], stats['subStat2Key'], stats['subStat3Key'], stats['subStat4Key'])
        print(excelStr)
        pyperclip.copy(excelStr)