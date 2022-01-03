
from time import sleep
import cv2
import numpy as np
import pyautogui
import win32gui, win32api, win32con

def click(x,y):
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)

def screenshot(window_title=None):
    if window_title:
        hwnd = win32gui.FindWindow(None, window_title)
        if hwnd:
            win32gui.SetForegroundWindow(hwnd)
            global x, y, x1, y1
            x, y, x1, y1 = win32gui.GetClientRect(hwnd)
            x, y = win32gui.ClientToScreen(hwnd, (x, y))
            x1, y1 = win32gui.ClientToScreen(hwnd, (x1 - x, y1 - y))
            im = pyautogui.screenshot(region=(x, y, x1, y1))
            return im
        else:
            print('Window not found!')
    else:
        im = pyautogui.screenshot()
        return im

def search():
    im = screenshot('Skype')
    if im:
        im.save(r'C:\\users\\aneurin\\Desktop\\Capcha\\image.png')

    img_rgb = cv2.imread('image.png')
    template = cv2.imread('arrow.png')
    w, h = template.shape[:-1]

    res = cv2.matchTemplate(img_rgb, template, cv2.TM_CCOEFF_NORMED)
    threshold = .6
    loc = np.where(res >= threshold)
    global saved
    distance = 4
    if len(loc[0]) > 0:
        saved = [loc[1][0],loc[0][0]]
        cv2.rectangle(img_rgb, saved, (saved[0] + w, saved[1] + h), (0, 0, 255), 2)

    for pt in zip(*loc[::-1]):  # Switch collumns and rows
        if not(pt[0] > saved[0]-distance and pt[0] < saved[0]+distance and pt[1] > saved[1]-distance and pt[1] < saved[1]+distance):
            cv2.rectangle(img_rgb, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 2)
            print('found')
            cv2.imwrite('result.png', img_rgb)
            return 1
        saved = pt
    cv2.imwrite('result.png', img_rgb)
    return 0

global done
done = 0
while done < 10:
    while search() != 1:
        print("not found")
        click(x+320,y+320)
        sleep(.35)

    click(x+260,y+430)
    done = done + 1
    print(done)
    sleep(1)