from pywinauto import application
from pywinauto import keyboard
import time

app = application.Application(backend='uia')
app.start("C:\Motion\Motion_E\Motion_E.exe")


def login(title, id):
    login_window = app.window(title=title)
    login_window.child_window(auto_id=id).click()


time.sleep(1)
if app.connect(path="C:\Motion\Motion_E\Motion_E.exe"):
    print('connect pass')
    login_window = app.window(title="로그인")
    motion_window = app.window(title='모션.ver[1.0.0.127]ipc://Motion_E-64bit')

    if login_window.exists():
        login('로그인', 'btnLogin')
        print("login pass")
else:
    print("connect fail")

time.sleep(5)
print("-------------------")

motion_window['닫기'].click()
