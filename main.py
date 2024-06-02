import multiprocessing
import ctypes
import sys
from func.dto.dto import DashboardDto
from func.start.motion_starter import *
from func.dashboard.dashboard import *
from func.publicFunc.public_func import *

MAX_RETRY = 3


class ProcessFunc():
    rad_box = None
    retries = 0

    def main_process_func(start_sub_process_event, sub_process_done_event):

        win32_app = application.Application(backend='win32')
        motion_app = application.Application(backend='uia')
        MotionStarter.app_connect(win32_app, motion_app)
        motion_window = motion_app.window(
            title=MotionStarter.version_search('모션.ver'))
        dto = DashboardDto(motion_window, motion_app, "2351", "01074417631",
                           start_sub_process_event, sub_process_done_event, "btnSave", "")
        start_sub_process_event.set()
        sub_process_done_event.wait()

        DashBoard.receipt_check(motion_window, "0000002351")

    def sub_process_func(start_sub_process_event, sub_process_done_event, window_auto_id, btn_auto_id):
        start_sub_process_event.wait()
        win32_app = application.Application(backend='win32')
        motion_app = application.Application(backend='uia')
        MotionStarter.app_connect(win32_app, motion_app)
        sub_process_done_event.set()
        start_sub_process_event.clear()

        while ProcessFunc.retries <= MAX_RETRY:
            try:
                start_sub_process_event.wait()
                registration_window = motion_app.window(
                    title=MotionStarter.version_search('고객등록'))
                ProcessFunc.rad_box = registration_window.child_window(
                    auto_id=window_auto_id, first_only=True)
                ProcessFunc.rad_box.wait(wait_for='exists enabled', timeout=30)
                rad_btn = ProcessFunc.rad_box.child_window(auto_id=btn_auto_id)
                rad_btn.click()
                sub_process_done_event.set()
                start_sub_process_event.clear()
            except Exception as e:
                print("서브 모듈 동작 실패 : ", e)
                retries += 1
                continue


if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, ' '.join(sys.argv), None, 1)
    sys.exit()


if __name__ == "__main__":

    start_sub_process_event = multiprocessing.Event()
    sub_process_done_event = multiprocessing.Event()

    main_process = multiprocessing.Process(
        target=ProcessFunc.main_process_func, args=(start_sub_process_event, sub_process_done_event))
    sub_process = multiprocessing.Process(
        target=ProcessFunc.sub_process_func, args=(start_sub_process_event, sub_process_done_event, "RadMessageBox", "radButton1"))

    main_process.start()
    sub_process.start()

    main_process.join()
    sub_process.terminate()
