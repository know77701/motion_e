from pywinauto import application
from pywinauto import Desktop
from pywinauto import keyboard
import time


app = application.Application(backend='uia')

class motionStarter:
    def login(title, id):
        login_window = app.window(title=title)
        login_window.child_window(auto_id=id).click()
    def appConnect():
        try:
            app.connect(path="C:\\Motion\\Motion_E\\Motion_E.exe")

        except application.ProcessNotFoundError:
            app.start("C:\\Motion\\Motion_E\\Motion_E.exe")
            time.sleep(3)
            motionStarter.login('로그인', 'btnLogin')
            print("로그인 성공")

    def motionVersionSearch():
        windows = Desktop(backend="uia").windows()

        for window in windows:
            try:
                if('모션.ver' in window.window_text()):
                    MotionTitle = window.window_text()
                    motion_window = app.window(title=MotionTitle)
                    print('app 연결 성공')
            except Exception as e:
                print('버전 찾기 실패', e)
        try:
            if app.window(title=MotionTitle):

        except Exception as e:
            print("app 연결 실패: ",e)
            

print("-------------------")


class dashBoard():
    def searchUser(searchName):
        motion_window.child_window(auto_id="srch-val",  control_type="Edit").type_keys(searchName)
        motion_window.child_window(title="검색", control_type="Button").click()
    def reserve():
        try:
            dashBoard.searchUser('김지헌')
            time.sleep(1)
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
            print("예약 성공: ", e)
        except Exception as e:
            keyboard.send_keys('{F5}')
            print("예약 실패: ", e)
    
    def receipt(searchName):
        try:
            dashBoard.searchUser('김지헌')
            time.sleep(1)
            motion_window.child_window(title="접수하기", control_type="Button", found_index=0).click()
            time.sleep(5)
            motion_window.child_window(
                auto_id="265281",  control_type="Edit").type_keys('접수메모')
            motion_window.child_window(
                auto_id="265210",  control_type="Edit").type_keys('메모하나')
            time.sleep(1)
            motion_window.child_window(auto_id="btnAcpt", control_type="Button").click()
            print("접수 성공")
        except Exception as e:
            keyboard.send_keys('{F5}')
            print("접수 실패: ", e)

class notice:
    def noticeCreate(value):
        try:
            print('공지사항 생성 성공')
            motion_window.child_window(auto_id="notice-content",  control_type="Edit").type_keys(value)
            keyboard.send_keys('{ENTER}')
        except Exception as e:
            print('공지사항 생성 실패: ', e)
            keyboard.send_keys('{F5}')

    def noticeDelete():
        try:
            motion_window.child_window(title="닫기", control_type="Button", found_index=0).click()
            time.sleep(1)
            
            motion_window.child_window(auto_id="radButton1", framework_id="WinForm").click()
            print('공지사항 삭제 성공')
        except Exception as e:
            print('공지사항 삭제 실패: ', e)



MotionTitle= ""
motionStarter.appConnect()
time.sleep(5)


motion_window = ""
time.sleep(5)
dashBoard.receipt('김지헌')

# dashboardAction.reserve('김지헌')
# time.sleep(5)
# notice.noticeCreate('테스트')
# time.sleep(3)
# notice.noticeDelete()