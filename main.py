import multiprocessing
import ctypes
import sys
from func.dto.dto import DashboardDto
from func.start.motion_starter import *
from func.dashboard.dashboard import *
from func.publicFunc.public_func import *

MAX_RETRY = 3

if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, ' '.join(sys.argv), None, 1)
    sys.exit()


class ProcessFunc():
    rad_box = None
    retries = 0

    def main_process_func(start_sub_process_event, sub_process_done_event):

        win32_app = application.Application(backend='win32')
        motion_app = application.Application(backend='uia')
        MotionStarter.app_connect(win32_app, motion_app)
        motion_window = motion_app.window(
            title=MotionStarter.version_search('모션.ver'))

        dto = DashboardDto(motion_window, motion_app, "자동화QA", "01074417631",
                           start_sub_process_event, sub_process_done_event, "")
        dto.start_sub_process_event.set()
        dto.sub_process_done_event.wait()

        DashBoard.save_receipt_popup(dto)

    def sub_process_func(start_sub_process_event, sub_process_done_event):
        start_sub_process_event.wait()
        win32_app = application.Application(backend='win32')
        motion_app = application.Application(backend='uia')
        MotionStarter.app_connect(win32_app, motion_app)
        sub_process_done_event.set()
        start_sub_process_event.clear()
        start_sub_process_event.wait()

        # while ProcessFunc.retries <= MAX_RETRY:
        try:
            registration_window = motion_app.window(
                title=MotionStarter.version_search('고객등록'))
            child_window = registration_window.children()
            close_btn = None
            for child in child_window:
                child_list = child.children()
                for items in child_list:
                    for item in items.children():
                        if item.element_info.control_type == "Button" and item.element_info.name == "확인":
                            close_btn = item
                            if item.element_info.name == "입력한 고객과 동일한 고객 정보(이름, 휴대폰번호)가 존재합니다" or item.element_info.name == "이름을 입력하세요.":
                                print(item)
            close_btn.click()
            sub_process_done_event.set()
            start_sub_process_event.clear()

            start_sub_process_event.wait()
        except Exception as e:
            print("서브 모듈 동작 실패 : ", e)
            retries += 1
            # continue


if __name__ == "__main__":

    start_sub_process_event = multiprocessing.Event()
    sub_process_done_event = multiprocessing.Event()

    main_process = multiprocessing.Process(
        target=ProcessFunc.main_process_func, args=(start_sub_process_event, sub_process_done_event))
    sub_process = multiprocessing.Process(
        target=ProcessFunc.sub_process_func, args=(start_sub_process_event, sub_process_done_event))

    main_process.start()
    sub_process.start()

    main_process.join()
    sub_process.terminate()
