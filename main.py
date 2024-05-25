from pywinauto import application, Desktop, keyboard
import time
import ctypes
import sys
import random


MAX_RETRY = 2


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


class MotionStarter:
    def version_search(search_title):
        windows = Desktop(backend="uia").windows()

        for window in windows:
            try:
                window_text = window.window_text()
                if search_title in window_text:
                    motion_title = window_text
                    return motion_title
            except Exception as e:
                print('버전 찾기 실패', e)

    @staticmethod
    def login_click(title, id):
        login_window = win32_app.window(title=title)
        login_window.child_window(auto_id=id).click()

    @staticmethod
    def app_title_connect(title, btnName):
        win32_app.connect(path="C:\\Motion\\Motion_E\\Motion_E.exe")
        MotionStarter.login_click(title, btnName)
        win32_app.kill()
        time.sleep(5)
        motion_app.connect(path="C:\\Motion\\Motion_E\\Motion_E.exe")

    @staticmethod
    def app_connect(retries=0):
        try:
            if MotionStarter.version_search('모션.ver'):
                motion_app.connect(
                    path="C:\\Motion\\Motion_E\\Motion_E.exe")
                print('기존 앱 연결')
            elif MotionStarter.version_search('로그인'):
                MotionStarter.app_title_connect('로그인', 'btnLogin')
                print('로그인 성공')
            else:
                win32_app.start("C:\\Motion\\Motion_E\\Motion_E.exe")
                time.sleep(2)
                MotionStarter.login_click('로그인', 'btnLogin')
                time.sleep(3)
                motion_app.connect(
                    path="C:\\Motion\\Motion_E\\Motion_E.exe")

        except application.ProcessNotFoundError as e:
            print("앱 찾기 실패 :", e)
            if retries < MAX_RETRY:
                retries += 1
                print(f"재시도 횟수: {retries}")
                win32_app.start("C:\\Motion\\Motion_E\\Motion_E.exe")
                MotionStarter.app_connect(retries)
            else:
                print("최대 재시도 횟수에 도달했습니다. 프로그램을 종료합니다.")
        except application.AppStartError:
            print("앱 미설치 또는 앱 미존재")
        except Exception as e:
            print("알수없는 에러 발생")


class DashBoard():
    @staticmethod
    def search_user(search_name):
        serach_window = motion_window.child_window(
            auto_id="srch-val",  control_type="Edit")
        serach_window.set_edit_text("")
        time.sleep(1)
        serach_window.set_edit_text(search_name)
        time.sleep(1)
        motion_window.child_window(
            title="검색", control_type="Button").click()

    def text_edit_popup(user_name, phone):
        DashBoard.popup_view(user_name)
        registration_window = motion_app.window(
            title=MotionStarter.version_search('고객등록'))
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
        receipt_window = motion_app.window(
            title=MotionStarter.version_search('접수'))

    def save_reserve_popup(serach_name, phone_number):
        DashBoard.text_edit_popup(serach_name, phone_number)

    def popup_user_save(serach_name, phone_number):
        try:
            DashBoard.text_edit_popup(serach_name, phone_number)

        except Exception as e:
            print('저장 실패 : ', e)
            time.sleep(1)
            keyboard.send_keys('{F5}')

    # def popup_user_save(serach_name):
    #     try:
    #         DashBoard.popup_view(serach_name)
    #         DashBoard.text_edit_popup(serach_name)

    #         new_process(registration_window, '고객등록',
    #                     'radButton1', 'btnSave')

    #     except Exception as e:
    #         print('저장 실패 : ', e)
    #         keyboard.send_keys('{F5}')

    def popup_view(serach_name):
        DashBoard.search_user(serach_name)
        register_btn = motion_window.child_window(
            title="환자 등록 후 예약", control_type="Button")
        register_btn.wait(wait_for='exists enabled', timeout=30)
        register_btn.click()

    def receipt_card():
        print("")

# 관리자 권한

    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, ' '.join(sys.argv), None, 1)
        sys.exit()


if __name__ == "__main__":
    win32_app = application.Application(backend='win32')
    motion_app = application.Application(backend='uia')

    MotionStarter.app_connect()
    motion_window = motion_app.window(
        title=MotionStarter.version_search('모션.ver'))

    # sub_process_event = multiprocessing.Event()
    # start_sub_process_event = multiprocessing.Event()
    # done_sub_process_event = multiprocessing.Event()
    # main_process = multiprocessing.Process(
    #     target=main, args=(start_sub_process_event,))
    # sub_process = multiprocessing.Process(target=process_func, args=())
    # main_process.start()
    # sub_process.start()
    # main_process.join()
