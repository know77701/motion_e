from PIL import ImageGrab
from functools import partial
from pywinauto import application, Desktop, keyboard, findwindows
import time
import ctypes
import sys
import os
import multiprocessing


MAX_RETRY = 3
screenshot_save_dir = "fail" 

def winodw_screen_shot(save_file_name):
    ImageGrab.grab = partial(ImageGrab.grab, all_screens=True)
    screenshot_path = os.path.join(screenshot_save_dir, save_file_name)
    save_image = ImageGrab.grab()
    save_image.save(screenshot_path)
    
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def process_func(window_title, button_auto_id):
    try:
        quit_event = multiprocessing.Event()
        MotionStarter.appConnect()
        while not quit_event.is_set():
            try:
                popup = MotionApp.window(
                    title=MotionStarter.VersionSearch(window_title))
                popup_button = popup.child_window(
                    auto_id=button_auto_id, control_type="Button")
                popup_button.wait('visible')
                popup_button.click()
                quit_event.set()
            except findwindows.ElementNotFoundError as e:
                print(f"버튼 미존재: {e}")
                pass
            except application.timings.TimeoutError as e:
                print(f"시간초과: {e}")
                pass
            time.sleep(1)
    except Exception as e:
        print("프로세스 실패", e)


def new_process(window_name, popup_name, btn_auto_id, click_btn_id):
    if __name__ == '__main__':
        quit_event = multiprocessing.Event()
        process1 = multiprocessing.Process(
            target=process_func, args=(popup_name, btn_auto_id, quit_event))
        process1.start()
        window_name.child_window(auto_id=click_btn_id).click()
        quit_event.set()
        process1.join()


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
        login_window = win32_app.window(title=title)
        login_window.child_window(auto_id=id).click()

    @staticmethod
    def appTitleAction(title, btnName):
        win32_app.connect(path="C:\\Motion\\Motion_E\\Motion_E.exe")
        MotionStarter.loginClick(title, btnName)
        win32_app.kill()
        time.sleep(5)
        MotionApp.connect(path="C:\\Motion\\Motion_E\\Motion_E.exe")

    @staticmethod
    def appConnect(retries=0):
        try:
            if MotionStarter.VersionSearch('모션.ver'):
                MotionApp.connect(
                    path="C:\\Motion\\Motion_E\\Motion_E.exe")
                print('기존 앱 연결')
            elif MotionStarter.VersionSearch('로그인'):
                MotionStarter.appTitleAction('로그인', 'btnLogin')
                print('로그인 성공')
            else:
                win32_app.start("C:\\Motion\\Motion_E\\Motion_E.exe")
                time.sleep(2)
                MotionStarter.loginClick('로그인', 'btnLogin')
                time.sleep(3)
                MotionApp.connect(
                    path="C:\\Motion\\Motion_E\\Motion_E.exe")

        except application.ProcessNotFoundError as e:
            print("앱 찾기 실패 :", e)
            if retries < MAX_RETRY:
                retries += 1
                print(f"재시도 횟수: {retries}")
                win32_app.start("C:\\Motion\\Motion_E\\Motion_E.exe")
                MotionStarter.appConnect(retries)
            else:
                print("최대 재시도 횟수에 도달했습니다. 프로그램을 종료합니다.")
        except application.AppStartError:
            print("앱 미설치 또는 앱 미존재")


class DashBoard:
    @staticmethod
    def searchUser(searchName):
        serach_window = motion_window.child_window(
            auto_id="srch-val",  control_type="Edit")
        print("테스트")
        serach_window.set_edit_text("")
        print("테스트")
        time.sleep(3)
        serach_window.set_edit_text(searchName)
        time.sleep(1)
        motion_window.child_window(
            title="검색", control_type="Button").click()

    @staticmethod
    def comboBox(count, index):
        for _ in range(count):
            motion_window.child_window(
                control_type="ComboBox", found_index=index).type_keys("{DOWN}" * 5)

    # 예약
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

    # 접수
    @staticmethod
    def receipt(receiptName):
        try:
            DashBoard.searchUser(receiptName)
            time.sleep(1)
            motion_window.child_window(
                title="접수하기", control_type="Button", found_index=0).click()

            time.sleep(1)
            receipt_window = MotionApp.window(
                title=MotionStarter.VersionSearch('접수'))
            new_process(receipt_window, '접수', 'radButton1', 'btnAcpt')
            print("접수 성공")
        except Exception as e:
            keyboard.send_keys('{F5}')
            print("접수 실패: ", e)

    def text_edit_popup(user_name, phone):
        DashBoard.popup_view(user_name)
        edit_window = registration_window.child_window(
            control_type="Edit", auto_id="txtPat_Nm")
        edit_window.set_edit_text(user_name)
        if phone:
            mobile_edit2 = registration_window.child_window(
                control_type="Edit", auto_id="txtMobile_No2")
            mobile_edit3 = registration_window.child_window(
                control_type="Edit", auto_id="txtMobile_No3")
            mobile_edit2.set_edit_text(phone[3:7])
            mobile_edit3.set_edit_text(phone[7:11])

    def save_receipt_popup(serach_name, phone_number):

        DashBoard.text_edit_popup(serach_name, phone_number)
        new_process(registration_window, '고객등록',
                    'radButton1', 'btnSaveRsrv')
        receipt_window = MotionApp.window(
            title=MotionStarter.VersionSearch('접수'))
        new_process(receipt_window, '접수', 'radButton1', 'btnAcpt')

    def save_reserve_popup(serach_name, phone_number):
        DashBoard.text_edit_popup(serach_name, phone_number)
        new_process(registration_window, '고객등록',
                    'radButton1', 'btnSaveAcpt')

    def popup_user_save(serach_name, phone_number):
        try:
            DashBoard.text_edit_popup(serach_name, phone_number)
            if __name__ == '__main__':
                process1 = multiprocessing.Process(
                    target=process_func, args=('고객등록', 'radButton1'))
                process1.start()
                registration_window.child_window(auto_id='btnSave').click()

                process1.terminate()
                process1.join()

        except Exception as e:
            print('저장 실패 : ', e)
            time.sleep(1)
            keyboard.send_keys('{F5}')

    # def popup_user_save(serach_name):
    #     try:
    #         DashBoard.popup_view(serach_name)
    #         DashBoard.text_edit_popup(serach_name)

            # new_process(registration_window, '고객등록',
    #                     'radButton1', 'btnSave')

    #     except Exception as e:
    #         print('저장 실패 : ', e)
    #         keyboard.send_keys('{F5}')

    def popup_view(serach_name):
        DashBoard.searchUser(serach_name)
        register_btn = motion_window.child_window(
            title="환자 등록 후 예약", control_type="Button")
        register_btn.wait(wait_for='exists enabled', timeout=30)
        register_btn.click()


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

def modal_btn_click(index):
    save_pen_chart_window = chart_window.child_window(
        auto_id="PenChartAddFrom")
    for _ in range(index):
        save_popup = save_pen_chart_window.child_window(
            auto_id="RadMessageBox")
        save_popup_btn = save_popup.child_window(auto_id="radButton1")
        save_popup_btn.click()
        time.sleep(1)



# 관리자 권한
if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, ' '.join(sys.argv), None, 1)
    sys.exit()

win32_app = application.Application(backend='win32')
MotionApp = application.Application(backend='uia')
MotionStarter.appConnect()
motion_window = MotionApp.window(
    title=MotionStarter.VersionSearch('모션.ver'))
def main():
    print("테스트")
if __name__ == '__main__':
    quit_event = multiprocessing.Event()
    main_process = multiprocessing.Process(target=main,args=())
    sub_process = multiprocessing.Process(target=modal_btn_click,args=3)
    main_process.start()
    
    
    
    
