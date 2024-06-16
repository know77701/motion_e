from pywinauto import application, Desktop, keyboard
import time
import ctypes
import sys
import multiprocessing

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

    def popup_view(search_name):
        DashBoard.search_user(search_name)
        register_btn = motion_window.child_window(
            title="환자 등록 후 예약", control_type="Button")
        register_btn.wait(wait_for='exists enabled', timeout=30)
        register_btn.click()

    def text_edit_popup(serach_name, phone_number, start_sub_process_event, sub_process_done_event, btn_auto_id):
        DashBoard.popup_view(serach_name)
        registration_window = motion_app.window(
            title=MotionStarter.version_search('고객등록'))
        edit_window = registration_window.child_window(
            control_type="Edit", auto_id="txtPat_Nm")
        edit_window.set_edit_text(serach_name)
        if phone_number:
            mobile_edit2 = registration_window.child_window(
                control_type="Edit", auto_id="txtMobile_No2")
            mobile_edit3 = registration_window.child_window(
                control_type="Edit", auto_id="txtMobile_No3")
            match len(phone_number):
                case 13:
                    mobile_edit2.set_edit_text(phone_number[4:8])
                    mobile_edit3.set_edit_text(phone_number[10:13])
                case 11:
                    mobile_edit2.set_edit_text(phone_number[3:7])
                    mobile_edit3.set_edit_text(phone_number[7:11])
                case 8:
                    mobile_edit2.set_edit_text(phone_number[1:4])
                    mobile_edit3.set_edit_text(phone_number[4:8])
        save_btn = registration_window.child_window(
            control_type="Button", auto_id=btn_auto_id)
        start_sub_process_event.set()
        save_btn.click()
        sub_process_done_event.wait()

    def save_receipt_popup(serach_name, phone_number, start_sub_process_event, sub_process_done_event, btn_auto_id):
        try:
            DashBoard.text_edit_popup(
                serach_name, phone_number, start_sub_process_event, sub_process_done_event, btn_auto_id)
            receipt_window = motion_app.window(
                title=MotionStarter.version_search('접수'))
            edit_field = receipt_window.child_window(auto_id="radPanel6")
            receipt_memo = edit_field.child_window(control_type="Edit")
            user_memo = edit_field.child_window(control_type="Edit")
            receipt_memo.set_text("테스트")
            user_memo.set_text("테스트")
            receipt_btn = receipt_window.child_window(
                auto_id="btnAcpt", control_type="Button")
            start_sub_process_event.set()
            receipt_btn.click()
            sub_process_done_event.wait()

            web_window = motion_app.child_winodw(title="Motion E web")
            acpt_list = web_window.child_window(
                control_type="List", auto_id="acpt-list")
            test = acpt_list.children()
            print(test)

        except Exception as e:
            if MotionStarter.version_search('고객등록'):
                registration_window = motion_app.window(
                    title=MotionStarter.version_search('고객등록'))
                close_btn = registration_window.child_window(
                    title="닫기", control_type="Button")
                close_btn.click()
                print(e)
            elif MotionStarter.version_search('접수'):
                receipt = motion_app.window(
                    title=MotionStarter.version_search('접수'))
                close_btn = receipt.child_window(
                    title="닫기", control_type="Button")
                close_btn.click()
                print(e)

    def save_reserve_popup(serach_name, phone_number, start_sub_process_event, sub_process_done_event, btn_auto_id):
        try:
            DashBoard.text_edit_popup(
                serach_name, phone_number, start_sub_process_event, sub_process_done_event, btn_auto_id)
            time.sleep(2)

        except:
            if MotionStarter.version_search('고객등록'):
                receipt_window = motion_app.window(
                    title=MotionStarter.version_search('고객등록'))
                close_btn = receipt_window.child_window(
                    title="닫기", control_type="Button")
                close_btn.click()
            else:
                top_menu = motion_app.child_window(auto_id="pnTop")
                dashboard_menu = top_menu.child_window(title="Dashboard")
                dashboard_menu.click_input()

    def user_save(serach_name, phone_number, start_sub_process_event, sub_process_done_event, btn_auto_id):
        try:
            DashBoard.text_edit_popup(
                serach_name, phone_number, start_sub_process_event, sub_process_done_event, btn_auto_id)
            time.sleep(3)

        except Exception as e:
            print('저장 실패 : ', e)
            if MotionStarter.version_search('고객등록'):
                registration_window = motion_app.window(
                    title=MotionStarter.version_search('고객등록'))
                close_btn = registration_window.child_window(
                    title="닫기", control_type="Button")
                close_btn.click()
            keyboard.send_keys('{F5}')

    def receipt_card():
        print("")

    def reserve_card():
        print("")


def main_process_func(start_sub_process_event, sub_process_done_event):
    # DashBoard.user_save("김지헌", "01074417631",
    #                     start_sub_process_event, sub_process_done_event, "btnSave")
    # # sub process unset
    # sub_process_done_event.clear()
    # start_sub_process_event.clear()
    # time(1)

    # DashBoard.save_reserve_popup("김지헌", "01074417631",
    #                              start_sub_process_event, sub_process_done_event, "btnSaveAcpt")
    # sub_process_done_event.clear()
    # start_sub_process_event.clear()
    # time(1)

    DashBoard.save_receipt_popup("김지헌", "01074417631",
                                 start_sub_process_event, sub_process_done_event, "btnSave")
    sub_process_done_event.clear()
    start_sub_process_event.clear()


def sub_process_func(start_sub_process_event, sub_process_done_event, window_auto_id, btn_auto_id):
    try:
        start_sub_process_event.wait()
        registration_window = motion_app.window(
            title=MotionStarter.version_search('고객등록'))
        while True:
            rad_box = registration_window.child_window(
                auto_id=window_auto_id)
            # rad_box.wait(wait_for='exists enabled', timeout=30)
            rad_btn = rad_box.child_window(auto_id=btn_auto_id)
            rad_btn.click()
            sub_process_done_event.set()
            start_sub_process_event.clear()
    except Exception as e:
        print("서브 모듈 동작 실패 : ", e)


if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, ' '.join(sys.argv), None, 1)
    sys.exit()
win32_app = application.Application(backend='win32')
motion_app = application.Application(backend='uia')
MotionStarter.app_connect()

motion_window = motion_app.window(title=MotionStarter.version_search('모션.ver'))

if __name__ == "__main__":

    start_sub_process_event = multiprocessing.Event()
    sub_process_done_event = multiprocessing.Event()

    main_process = multiprocessing.Process(
        target=main_process_func, args=(start_sub_process_event, sub_process_done_event))
    sub_process = multiprocessing.Process(
        target=sub_process_func, args=(start_sub_process_event, sub_process_done_event, "RadMessageBox", "radButton1"))

    main_process.start()
    sub_process.start()

    main_process.join()
    sub_process.terminate()