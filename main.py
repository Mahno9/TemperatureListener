import random

import win32api
import win32gui
import time
import ctypes
from ctypes import wintypes
import threading
import sys

import xlsxwriter


def GetClickPosition() -> ((int, int), (int, int)):
    state_left = win32api.GetKeyState(0x01)
    returnValues = []
    while True:
        a = win32api.GetKeyState(0x01)
        if a != state_left:  # Button state changed
            state_left = a
            point = win32api.GetCursorPos()
            if a < 0:
                returnValues.append(point)
            else:
                returnValues.append(point)
                break

        time.sleep(0.001)

    assert len(returnValues) == 2
    return returnValues[0], returnValues[1]


def FindWindowByClick():
    _, clickPos = GetClickPosition()
    return win32gui.WindowFromPoint(clickPos)


def GetChildsHWDNList(hwndParent: ctypes.wintypes.HWND):
    def EnumFunc(hwnd: ctypes.wintypes.HWND, resultList):
        resultList.append(hwnd)
        return True

    childsHWNDList = []
    win32gui.EnumChildWindows(hwndParent, EnumFunc, childsHWNDList)
    return childsHWNDList


def RequestLabelNumber(childsHWND) -> int:
    i = 0
    for childHWND in childsHWND:
        print("{}: Subwindow text: {}".format(i, win32gui.GetWindowText(childHWND)))
        i += 1
    return int(input("Type a number of the label you want listen to: "))


def GetParentWindowHWDN():
    while (True):
        print("Click to any window to find labes in it.")
        clickedWindowHWND = FindWindowByClick()
        print("Clicked widnow HWND: {}".format(hex(clickedWindowHWND)))
        print("Clicked widnow Title: {}".format(win32gui.GetWindowText(clickedWindowHWND)))

        answer = input("Is this proper window? Y[es]/N[o]")
        if any(answer in s for s in ['Yes', 'yes', 'Y', 'y', '', '\n']):
            return clickedWindowHWND


def StartLabelStatusListening(listenHWND, pollingInterval):
    data = []
    needStopThread = False

    def grabFunc():
        import time
        startTime = time.time()
        while not needStopThread:
            timePassed = time.time() - startTime
            value = win32gui.GetWindowText(listenHWND)
            data.append((timePassed, value))
            print(timePassed, value)
            time.sleep(pollingInterval)

    thread = threading.Thread(target=grabFunc)
    thread.start()
    sys.stdin.read(1)
    needStopThread = True
    thread.join()
    return data


def main():
    clickedWindowHWND = GetParentWindowHWDN()

    children = GetChildsHWDNList(clickedWindowHWND)
    if len(children) == 0:
        print("Unable to find any subwindows for this window :(")
        return

    listenLabelIdx = RequestLabelNumber(children)
    print("\nListen label number is {}".format(listenLabelIdx))

    listenHWND = children[listenLabelIdx]
    pollingInterval = input("Please specify polling interval in sec (default: 0.1): ")
    if pollingInterval == '':
        pollingInterval = 0.1
    else:
        pollingInterval = float(pollingInterval)

    data = StartLabelStatusListening(listenHWND, pollingInterval)
    DumpDataToXlsx(data)


def DumpDataToXlsx(data, outFilename=None):
    if len(data) == 0:
        print("Nothing to export.")
        return

    if outFilename is None:
        outFilename = 'Result.xlsx'

    workbook = xlsxwriter.Workbook(outFilename)
    worksheet = workbook.worksheets()[0] if len(workbook.worksheets()) > 0 else workbook.add_worksheet()
    for i in range(len(data[0])):
        worksheet.write_column(0, i, [''] * (len(data) + 1))
        columnData = list(zip(*data))[i]
        worksheet.write_column(0, i, columnData)
    workbook.close()


if __name__ == '__main__':
    main()
