from pywinauto import application
import time

app = application.Application(backend='win32')

def login(title, id):
    login_window = app.window(title=title)
    login_window.child_window(auto_id=id).click()


def appConnect():
    try:
        app.connect(path="C:\Motion\Motion_E\Motion_E.exe")
        print('기존 앱 연결')

    except application.ProcessNotFoundError:
        app.start("C:\Motion\Motion_E\Motion_E.exe")
        time.sleep(1)

        print("로그인 성공")


appConnect()
login('로그인', 'btnLogin')
motion_window = app.window(title='모션.ver[1.0.1.200]ipc://Motion_E-64bit')
print('모션 연결 성공')
print("-------------------")
# time.sleep(3)

