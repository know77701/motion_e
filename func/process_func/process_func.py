from pywinauto import application
from func.start.motion_starter import *
from func.dashboard.dashboard import *
from func.process_func.process_func import *

MAX_RETRY = 3


class ProcessFunc():
    rad_box = None

    def main_process_func(start_sub_process_event, sub_process_done_event, motion_app, motion_window):
        DashBoard.user_save("자동화체크1", "01074417631",
                            start_sub_process_event, sub_process_done_event, "btnSave", motion_app, motion_window)
        # sub process unset
        sub_process_done_event.clear()
        start_sub_process_event.clear()
        time.sleep(1)

        DashBoard.save_reserve_popup("자동화체크2", "01074417631",
                                     start_sub_process_event, sub_process_done_event, "btnSaveRsrv", motion_app, motion_window)
        sub_process_done_event.clear()
        start_sub_process_event.clear()
        time.sleep(1)

        DashBoard.save_receipt_popup("자동화체크3", "01074417631",
                                     start_sub_process_event, sub_process_done_event, "btnSaveAcpt", motion_app, motion_window)
        sub_process_done_event.clear()
        start_sub_process_event.clear()

    def sub_process_func(start_sub_process_event, sub_process_done_event, window_auto_id, btn_auto_id, motion_app):
        retries = 0
        while retries <= MAX_RETRY:
            try:
                start_sub_process_event.wait()
                registration_window = motion_app.window(
                    title=MotionStarter.version_search('고객등록'))
                ProcessFunc.rad_box = registration_window.child_window(
                    auto_id=window_auto_id)
                ProcessFunc.rad_box.wait(wait_for='exists enabled', timeout=30)
                rad_btn = ProcessFunc.rad_box.child_window(auto_id=btn_auto_id)
                rad_btn.click()
                sub_process_done_event.set()
                start_sub_process_event.clear()
            except Exception as e:
                print("서브 모듈 동작 실패 : ", e)
                retries += 1
                continue
