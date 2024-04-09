from pywinauto import application
from pywinauto import Desktop
import time


app = application.Application(backend='uia')


def login(title, id):
    login_window = app.window(title=title)
    login_window.child_window(auto_id=id).click()


def appConnect():
    try:
        app.connect(path="C:\\Motion\\Motion_E\\Motion_E.exe")
        print('기존 앱 연결')

    except application.ProcessNotFoundError:

        app.start("C:\\Motion\\Motion_E\\Motion_E.exe")
        windows = Desktop(backend="uia").windows()
        time.sleep(3)

        login('로그인', 'btnLogin')
        time.sleep(3)
        print("로그인 성공")


appConnect()


windows = Desktop(backend="uia").windows()
for window in windows:
    if('모션.ver' in window.window_text()):
        MotionTitle = window.window_text()
        

motion_window = app.window(title=MotionTitle)
print(motion_window)
print('모션 연결 성공')
print("-------------------")

def dashboardreserve(searchName):
    print('TEST')
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

time.sleep(5)
dashboardreserve('김지헌')
motion_window.child_window(
        AutomationId="srch-val",  control_type="Edit").type_keys('테스트')
motion_window.child_window(auto_id="notice-content").click()
motion_window.child_window(
        auto_id="notice-content",  control_type="Edit").type_keys("공지 테스트")
