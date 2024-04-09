from pywinauto import application
from pywinauto import Desktop
from pywinauto import keyboard
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
        time.sleep(3)
        login('로그인', 'btnLogin')
        time.sleep(3)
        print("로그인 성공")


appConnect()

MotionTitle= {}
windows = Desktop(backend="uia").windows()
for window in windows:
    if('모션.ver' in window.window_text()):
        MotionTitle = window.window_text()
        

motion_window = app.window(title=MotionTitle)
print('모션 연결 성공')
print("-------------------")

def dashboardreserve(searchName):
    motion_window.child_window(
        auto_id="srch-val",  control_type="Edit").type_keys(searchName)
    motion_window.child_window(title="검색", control_type="Button").click()
    motion_window.child_window(
        title="예약하기", control_type="Button", found_index=1).click()
    time.sleep(1)
    motion_window.child_window(title="오늘", control_type="Button").click()
    # motion_window.child_window(
    #     title="시간", control_type="ComboBox").click_input()

    # motion_window.child_window(
    #     title="20", control_type="ListItem").click_input()

    # motion_window.child_window(
    #     control_type="ComboBox", found_index=1).select()
    # motion_window.child_window(title="20", control_type="ListItem").click()

# dashboardreserve('김지헌')
# motion_window.child_window(
#         auto_id="srch-val",  control_type="Edit").type_keys('테스트')

class notice:
    @staticmethod
    def noticeCreate(value):
        try:
            print('공지사항 생성 성공')
            motion_window.child_window(auto_id="notice-content",  control_type="Edit").type_keys(value)
            keyboard.send_keys('{ENTER}')
        except Exception as e:
            print('공지사항 생성 실패:', e)

    def noticeDelete():
        try:
            motion_window.child_window(title="닫기", control_type="Button", found_index=0).click()
            time.sleep(1)
            
            print(motion_window.child_window(title="radButton1"))
            motion_window.child_window(auto_id="radButton1", control_type="Button").click()
            print('공지사항 삭제 성공')
        except Exception as e:
            print('공지사항 삭제 실패')


if notice.noticeCreate('테스트'):
    time.sleep(1)
    notice.noticeDelete()