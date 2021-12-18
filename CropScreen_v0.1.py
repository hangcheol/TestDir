import time
import pygetwindow as gw
import tkinter
from tkinter import *
import pywinauto
import win32gui, win32con
from PIL import Image, ImageGrab
import cv2
import time


# ----------------------------------------------------------------------------------------------------
# 마우스 드래그를 통한 이미지 크롭
def onMouse(event, x, y, flags, param):
    global isDragging, x0, y0, img
    if event == cv2.EVENT_LBUTTONDOWN:
        isDragging = True
        x0 = x
        y0 = y
    elif event == cv2.EVENT_MOUSEMOVE:
        if isDragging:
            img_draw = img.copy()
            cv2.rectangle(img_draw, (x0, y0), (x, y), red, 1)
            cv2.imshow(temptitle, img_draw)
    elif event == cv2.EVENT_LBUTTONUP:
        if isDragging:
            isDragging = False
            w = x - x0
            h = y - y0
            if w > 0 and h > 0:
                img_draw = img.copy()
                cv2.rectangle(img_draw, (x0, y0), (x, y), red, 1)
                cv2.imshow(temptitle, img_draw)
                roi = img[y0:y0 + h, x0:x0 + w]
                cv2.imwrite('crop_image.jpg', roi)
                # print(roi)
                cv2.destroyAllWindows()
            else:
                cv2.imshow('img', img)
                print('drag should start from left-top side')
# ----------------------------------------------------------------------------------------------------

isDragging = False
x0, y0, w, h = -1, -1, -1, -1
blue, red = (255, 0, 0), (0, 0, 255)

# 임시 활성화 이미지 경로
filename = 'test.jpg'
# 임시 활성화 이미지 창이름
temptitle = '-------temp_img-------'

# 현재 창 최소화(구동 웹 혹은 프로그램)
FrontWindow = win32gui.GetForegroundWindow()
win32gui.ShowWindow(FrontWindow, win32con.SW_MINIMIZE)
time.sleep(1)

# 캡처 및 저장
im = ImageGrab.grab() # 전체화면 캡쳐
im.save(filename, 'JPEG')

# 캡처 이미지 열기
img = cv2.imread(filename)
cv2.namedWindow(temptitle, cv2.WND_PROP_FULLSCREEN)
cv2.setWindowProperty(temptitle, cv2.WND_PROP_FULLSCREEN, cv2.WINDOW_FULLSCREEN)
cv2.imshow(temptitle, img)

# 캡처 이미지 최상단 활성화
totalwin = gw.getWindowsWithTitle(temptitle)
for win in totalwin:
    if temptitle == win32gui.GetWindowText(win._hWnd):
        handle = win._hWnd
        # print(handle)
        if not win.isActive:

            pywinauto.application.Application().connect(handle=handle).top_window().set_focus()
            win.activate
        break
    else:
        exit()

win32gui.IsIconic(handle) == 0

# 범위 지정 crop 및 저장
cv2.setMouseCallback(temptitle, onMouse)
cv2.waitKey()
cv2.destroyAllWindows()

# 최소화했던 구동프로그램 활성화
win32gui.ShowWindow(FrontWindow, win32con.SW_SHOWDEFAULT)
