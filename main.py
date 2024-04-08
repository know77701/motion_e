from pywinauto import application
import time
from pywinauto import keyboard


app = application.Application(backend='uia')


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
        login('로그인', 'btnLogin')
        print("로그인 성공")


appConnect()
motion_window = app.window(title='모션.ver[1.0.0.217]ipc://Motion_E-64bit')
print('모션 연결 성공')
print("-------------------")
# time.sleep(3)


def dashboardreserve(searchName):
    motion_window.child_window(
        auto_id="srch-val",  control_type="Edit").type_keys(searchName)
    motion_window.child_window(title="검색", control_type="Button").click()
    motion_window.child_window(
        title="예약하기", control_type="Button", found_index=1).click()
    time.sleep(1)
    motion_window.child_window(title="오늘", control_type="Button").click()
    motion_window.child_window(
        title="시간", control_type="ComboBox").click_input()

    motion_window.child_window(
        title="20", control_type="ListItem").click_input()

    motion_window.child_window(
        control_type="ComboBox", found_index=1).select()
    motion_window.child_window(title="20", control_type="ListItem").click()


dashboardreserve('김지헌')

time.sleep(2)
motion_window.child_window(auto_id="notice-content",
                           control_type="Edit").type_keys('테스트')
