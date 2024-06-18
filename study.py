from pywinauto import findwindows 
from pywinauto import application
import time
import os

current_path = os.getcwd()

print("현재위치 :" + current_path)
os.makedirs("fail")

# def mkdir(comare_folder_make):
    # comare_folder_make = os.mkdir("/fail")

# app_win32 = application.Application(backend='win32')
# app_uia = application.Application(backend='uia')

# # 이곳에 if문으로 프로세스가 돌아가는지 체크 후, 미실행 시 motion start

# # app_win32.start("C:\\Motion\\Motion_E\\Motion_E.exe")
# # print("실행 성공")

# connect_win = app_uia.connect(path="C:\\Motion\\Motion_E\\Motion_E.exe")
# print("커넥트 성공")

# login_window = connect_win.window(title='로그인')
# print("타이틀 확인 성공 ")

# time.sleep(10.0)

# login_window.child_window(title='로그인',control_type='Button').click()
# print("로그인 버튼 선택 성공")

















