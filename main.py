from pywinauto import application
import time

app = application.Application(backend='uia')


def login(title, id):
    login_window = app.window(title=title)
    login_window.child_window(auto_id=id).click()


try:
    app.connect(path="C:\Motion\Motion_E\Motion_E.exe")
    motion_window = app.window(title='모션.ver[1.0.0.127]ipc://Motion_E-64bit')
    print('기존 앱 연결')

except application.ProcessNotFoundError:
    app.start("C:\Motion\Motion_E\Motion_E.exe")
    login_window = app.window(title="로그인")
    time.sleep(1)
    if login_window.exists():
        login('로그인', 'btnLogin')
        print("로그인 성공")

time.sleep(5)
print('모션 연결 성공')
print("-------------------")
