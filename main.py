from pywinauto import application
from pywinauto import Desktop
from pywinauto import keyboard
from pywinauto import findwindows
import time


app = application.Application(backend='uia')


class MotionStarer():
    @staticmethod
    def login(title, id):
        login_window = app.window(title=title)
        login_window.child_window(auto_id=id).click()

    @staticmethod
    def appConnect():
        try:
            app.connect(path="C:\Motion\Motion_E\Motion_E.exe")
            print('기존 앱 연결')

        except application.ProcessNotFoundError:
            app.start("C:\Motion\Motion_E\Motion_E.exe")
            time.sleep(1)
            MotionStarer.login('로그인', 'btnLogin')
            print("로그인 성공")

    @staticmethod
    def VersionSearch():
        windows = Desktop(backend="uia").windows()

        for window in windows:
            try:
                if '모션.ver' in window.window_text():
                    MotionTitle = window.window_text()
                    return MotionTitle
            except Exception as e:
                print('버전 찾기 실패', e)


class DashBoard():
    @staticmethod
    def searchUser(searchName):
        motion_window.child_window(
            auto_id="srch-val",  control_type="Edit").type_keys(searchName)
        motion_window.child_window(title="검색", control_type="Button").click()

    @staticmethod
    def reserve(name):
        try:
            DashBoard.searchUser(name)
            time.sleep(1)
            motion_window.child_window(
                title="예약하기", control_type="Button", found_index=1).click()
            time.sleep(1)
            motion_window.child_window(
                title="오늘", control_type="Button").click()
            motion_window.child_window(
                title="시간", control_type="ComboBox").click_input()

            motion_window.child_window(
                title="20", control_type="ListItem").click_input()

            motion_window.child_window(
                control_type="ComboBox", found_index=1).select()
            motion_window.child_window(
                title="20", control_type="ListItem").click()
            print("예약 성공: ", e)
        except Exception as e:
            keyboard.send_keys('{F5}')
            print("예약 실패: ", e)

    @staticmethod
    def receipt(receiptName):
        try:
            DashBoard.searchUser(receiptName)
            time.sleep(1)
            motion_window.child_window(
                title="접수하기", control_type="Button", found_index=0).click()
            time.sleep(5)
            motion_window.child_window(
                auto_id="265281",  control_type="Edit").type_keys('접수메모')
            motion_window.child_window(
                auto_id="265210",  control_type="Edit").type_keys('메모하나')
            time.sleep(1)
            motion_window.child_window(
                auto_id="btnAcpt", control_type="Button").click()
            print("접수 성공")
        except Exception as e:
            keyboard.send_keys('{F5}')
            print("접수 실패: ", e)


class Notice:
    @staticmethod
    def noticeCreate(value):
        try:
            time.sleep(1)
            motion_window.child_window(
                auto_id="notice-content", control_type="Edit").type_keys(value)
            time.sleep(1)
            keyboard.send_keys('{ENTER}')
            print('공지사항 생성 성공')
        except Exception as e:
            print('공지사항 생성 실패: ', e)
            keyboard.send_keys('{F5}')

    def noticeDelete():
        try:
            motion_window.child_window(
                title="닫기", control_type="Button", found_index=0).click()
            time.sleep(1)

            motion_window.child_window(
                title="네", control_type="Button").click()
            print('공지사항 삭제 성공')
        except Exception as e:
            print('공지사항 삭제 실패: ', e)


MotionStarer.appConnect()
motion_window = app.window(title=MotionStarer.VersionSearch())
print('모션 연결 성공')
print("-------------------")

# time.sleep(10)
# DashBoard.receipt('김지헌')
# time.sleep(1)
# DashBoard.reserve('김지헌')
time.sleep(10)
Notice.noticeCreate('테스트')
time.sleep(5)
Notice.noticeDelete()

findwindows.find_element(title="네")
