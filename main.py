import win32api
import win32con
import win32gui
import time
import ctypes
from ctypes import wintypes, windll, byref


def GetClickPosition() -> ((int, int), (int, int)):
    state_left = win32api.GetKeyState(0x01)  # Left button down = 0 or 1. Button up = -127 or -128
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


def requestLabelNumber(childsHWND):
    i = 0
    for childHWND in childsHWND:
        print("{}: Subwindow text: {}".format(i, win32gui.GetWindowText(childHWND)))
        i += 1
    return input("Type a number of the label you want listen to: ")


def main():
    clickedWindowHWND = FindWindowByClick()
    print("Clicked widnow HWND: {}".format(hex(clickedWindowHWND)))
    print("Clicked widnow Title: {}".format(win32gui.GetWindowText(clickedWindowHWND)))

    childs = GetChildsHWDNList(clickedWindowHWND)
    if len(childs) == 0:
        print("Unable to find any subwindows for this window :(")
        return

    listenLabelNumber = requestLabelNumber(childs)
    print("Listen label number: {}".format(listenLabelNumber))


if __name__ == '__main__':
    main()
