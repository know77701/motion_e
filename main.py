from pywinauto import application
from pywinauto import Desktop

from pywinauto import keyboard
import time

app = application.Application(backend='win32')
MotionApp = application.Application(backend='uia')

MAX_RETRY = 3

class MotionStarter:
    def VersionSearch(value):
        windows = Desktop(backend="uia").windows()

        for window in windows:
            try:
                window_text = window.window_text()
                if value in window_text:
                    MotionTitle = window_text
                    return MotionTitle
            except Exception as e:
                print('버전 찾기 실패', e)

    @staticmethod
    def loginClick(title, id):
        login_window = app.window(title=title)
        login_window.child_window(auto_id=id).click()

    @staticmethod
    def appTitleAction(title, btnName):
        app.connect(path="C:\\Motion\\Motion_E\\Motion_E.exe")
        MotionStarter.loginClick(title, btnName)
        time.sleep(5)
        MotionApp.connect(path="C:\\Motion\\Motion_E\\Motion_E.exe")

    @staticmethod
    def appConnect(retries=0):
        try:
            if MotionStarter.VersionSearch('모션.ver'):
                MotionApp.connect(path="C:\\Motion\\Motion_E\\Motion_E.exe")
                print('기존 앱 연결')
            elif MotionStarter.VersionSearch('로그인'):
                print('로그인 테스트')
                MotionStarter.appTitleAction('로그인', 'btnLogin')
            else:
                app.start("C:\\Motion\\Motion_E\\Motion_E.exe")
                time.sleep(3)
                MotionStarter.loginClick('로그인', 'btnLogin')
                time.sleep(5)
                MotionApp.connect(path="C:\\Motion\\Motion_E\\Motion_E.exe")

        except application.ProcessNotFoundError as e:
            print("앱 찾기 실패 :", e)
            if retries < MAX_RETRY:
                retries += 1
                print(f"재시도 횟수: {retries}")
                app.start("C:\\Motion\\Motion_E\\Motion_E.exe")
                MotionStarter.appConnect(retries)
            else:
                print("최대 재시도 횟수에 도달했습니다. 프로그램을 종료합니다.")
        except application.AppStartError:
            print("앱 미설치 또는 앱 미존재")



class DashBoard():
    @staticmethod
    def searchUser(searchName):
        serach_window = motion_window.child_window(
            auto_id="srch-val",  control_type="Edit")
        serach_window.set_edit_text("")
        time.sleep(3)
        serach_window.set_edit_text(searchName)
        time.sleep(1)
        motion_window.child_window(title="검색", control_type="Button").click()

    @staticmethod
    def comboBox(count, index):
        for _ in range(count):
            motion_window.child_window(
                control_type="ComboBox", found_index=index).type_keys("{DOWN}")
            motion_window.child_window(
                control_type="ComboBox", found_index=index).type_keys("{DOWN}")
            motion_window.child_window(
                control_type="ComboBox", found_index=index).type_keys("{DOWN}")
            motion_window.child_window(
                control_type="ComboBox", found_index=index).type_keys("{DOWN}")
            motion_window.child_window(
                control_type="ComboBox", found_index=index).type_keys("{DOWN}")

    @staticmethod
    def reserve(name, index):
        try:
            DashBoard.searchUser(name)
            time.sleep(1)
            if motion_window.child_window(title=name, control_type="Text"):
                motion_window.child_window(
                    title="예약하기", control_type="Button", found_index=index).click()
                time.sleep(1)
                motion_window.child_window(
                    title="오늘", control_type="Button").click()
                DashBoard.comboBox(4, 0)
                DashBoard.comboBox(2, 1)

                motion_window.child_window(
                    title="예약", control_type="Button").click()
                time.sleep(3)
            keyboard.send_keys('{ENTER}')
            time.sleep(1)
            keyboard.send_keys('{F5}')
            print("예약 성공")
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

            # 접수창 control
            receipt_window = MotionApp.window(
                title=MotionStarter.VersionSearch('접수'))
            receipt_window.child_window(
                auto_id="btnAcpt", control_type="Button").click()
            time.sleep(1)
            print("test")
            message_box = receipt_window.window(auto_id="RadMessageBox")
            # 백그라운드에서 열어서 이미 찾을거같은데 이부분은 더 생각해보자
            if message_box :
                print("test2")
                message_box.window(
                    auto_id="radButton1", control_type="Button").click()
                time.sleep(1)
                print(message_box)
                message_box.window(auto_id="radButton1").click()
            else :
                print("접수되어있는 동일환자 미존재")
            time.sleep(1)
            keyboard.send_keys('{Enter}')
            time.sleep(1)
            keyboard.send_keys('{F5}')
            print("접수 성공")
        except Exception as e:
            keyboard.send_keys('{F5}')
            print("접수 실패: ", e)
    def registration():
        try:
            DashBoard.searchUser('2351')
            motion_window.child_window(
                title="환자 등록 후 예약", control_type="Button", found_index=0).click()
            registration_window = MotionApp.window(
                title=MotionStarter.VersionSearch('고객 등록'))
            edit_window = registration_window.child_window(control_type="Edit", auto_id="txtPat_Nm")
            mobile_edit2 = registration_window.child_window(control_type="Edit", auto_id="txtMobile_No2")
            mobile_edit3 = registration_window.child_window(control_type="Edit", auto_id="txtMobile_No3")

            edit_window.set_edit_text('QA테스트')
            mobile_edit2.set_edit_text('7441')
            mobile_edit3.set_edit_text('7631')
            registration_window.child_window(auto_id="btnSave", control_type="Button").click()
        except Exception as err:
            print('접수 실패 : ', err)

class Notice:
    def noticeCreate(value):
        try:
            time.sleep(1)
            motion_window.child_window(
                auto_id="notice-content", control_type="Edit").type_keys(value)
            time.sleep(1)
            keyboard.send_keys('{ENTER}')
            print('공지사항 생성 성공')
        except Exception as e:
            keyboard.send_keys('{F5}')
            print('공지사항 생성 실패: ', e)

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



time.sleep(1)

MotionStarter.appConnect()

motion_window = MotionApp.window(title=MotionStarter.VersionSearch('모션.ver'))



# 예약자 이름
# DashBoard.reserve('2351', 0)
# time.sleep(1)
# DashBoard.receipt('2351')
# DashBoard.registration()

registration_window = MotionApp.window(
    title=MotionStarter.VersionSearch('고객 등록'))
# edit_window = registration_window.child_window(control_type="Edit", auto_id="txtPat_Nm")
# mobile_edit2 = registration_window.child_window(control_type="Edit", auto_id="txtMobile_No2")
# mobile_edit3 = registration_window.child_window(control_type="Edit", auto_id="txtMobile_No3")

# edit_window.set_edit_text('QA테스트')
# mobile_edit2.set_edit_text('7441')
# mobile_edit3.set_edit_text('7631')
registration_window.child_window(auto_id="btnSave", control_type="Button").click()
registration_window.child_window(auto_id="radButton1", control_type="Button").click()

print(registration_window.child_window(auto_id="radButton1"))
registration_window.child_window(auto_id="radButton1", control_type="Button").click()
# if()


# 고객등록 > 고객등록 예약 > 고객 등록 접수 > 