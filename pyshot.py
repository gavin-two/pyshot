import os
import math
import time
## Windows API用
import ctypes
import win32gui
## Excel書き込み用
import xlsxwriter

## 保存ファイル用時間
sTime = time.time()

from PIL import ImageGrab

## 自分自身のファイル名
## sc_file_name = os.path.basename(__file__)
sc_file_name = "C:\\Windows\\py.exe"


## 定数
HWND_BOTTOM = 1
HWND_NOTOPMOST = -2
HWND_TOP = 0
HWND_TOPMOST = -1
SWP_HIDEWINDOW = 0x0080
SWP_NOREDRAW = 0x0008
SWP_SHOWWINDOW = 0x0040
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
GW_HWNDNEXT = 2

## -------------------------------------
## 現在アクティブなウィンドウ名を探す
## -------------------------------------
process_list = []
def callback(handle, _):
    process_list.append(win32gui.GetWindowText(handle))
win32gui.EnumWindows(callback, None)


## -------------------------------------
## 対象ウィンドウを設定
## -------------------------------------
hnd = win32gui.GetDesktopWindow()
if process_list:
    for process_name in process_list:
        if sc_file_name in process_name:
            print("入った")
            hSelfWnd = win32gui.FindWindow(None, process_name)
            hNextWnd = win32gui.GetWindow(hSelfWnd, GW_HWNDNEXT)
            win32gui.SetWindowPos(hSelfWnd, HWND_BOTTOM, 0, 0, 0, 0, (SWP_HIDEWINDOW))
            win32gui.SetWindowPos(hNextWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW | SWP_NOMOVE | SWP_NOSIZE);
            win32gui.BringWindowToTop(hNextWnd)
            hnd = win32gui.GetForegroundWindow()
            break
## else:
##     ## 見つからなかったら画面全体を取得
##     hnd = win32gui.GetDesktopWindow()

print(hnd)

## -------------------------------------
## ウィンドウサイズ取得
## -------------------------------------
x0, y0, x1, y1 = win32gui.GetWindowRect(hnd)

print(x0)
print(y0)
print(x1)
print(y1)

width = x1 - x0
height = y1 - y0

capSize = (width, height)

ImageGrab.grab(bbox=(x0, y0, x1, y1)).save("aaa.png")


## -------------------------------------
## 保存したスクショをExcelに貼り付け
## -------------------------------------

## 画像貼り付け用のExcelを作成
workbook = xlsxwriter.Workbook('capture.xlsx')
## シートを追加
worksheet = workbook.add_worksheet('aaa')
## 画像を貼り付け
worksheet.insert_image('A2', 'aaa.png', {'x_scale': 0.45, 'y_scale': 0.45})

## Excelを閉じる
workbook.close()

