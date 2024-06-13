from pywinauto import application, findwindows
from pywinauto.controls.hwndwrapper import HwndWrapper
from func.dto.dto import DashboardDto
from func.start.motion_starter import *
from func.dashboard.dashboard import *
from func.publicFunc.public_func import *
from func.chart.chart_func import *


class ProcessFunc():
    MAX_RETRY = 3
    rad_box = None
    retries = 0
    win32_value = 'win32'
    uia_value = 'uia'
    motion_value = '모션.ver'

    def main_process_func(start_sub_process_event, sub_process_done_event):

        win32_app = application.Application(backend=ProcessFunc.win32_value)
        motion_app = application.Application(backend=ProcessFunc.uia_value)
        MotionStarter.app_connect(win32_app, motion_app)
        motion_window = motion_app.window(
            title=MotionStarter.version_search(ProcessFunc.motion_value))

        dto = DashboardDto(motion_window, motion_app, "QA9", "01074417631",
                           start_sub_process_event, sub_process_done_event, "", "")

        # 서브프로세스 통신용
        dto.start_sub_process_event.set()

        # 서브프로세스 대기용
        dto.sub_process_done_event.wait()

        # 여기서부터 시작
        DashBoard.dashboard_starter(dto)
        ChartFunc.chart_starter()

    def sub_process_func(start_sub_process_event, sub_process_done_event):
        start_sub_process_event.wait()
        win32_app = application.Application(backend=ProcessFunc.win32_value)
        motion_app = application.Application(backend=ProcessFunc.uia_value)
        MotionStarter.app_connect(win32_app, motion_app)
        sub_process_done_event.set()

        start_sub_process_event.clear()

        while ProcessFunc.retries <= ProcessFunc.MAX_RETRY:
            try:
                start_sub_process_event.wait()
                time.sleep(2)
                ProcessFunc.rad_button_click()
                sub_process_done_event.set()
                start_sub_process_event.clear()
                ProcessFunc.retries = 0
                continue

            except Exception as e:
                print("서브 모듈 동작 실패 : ", e)
                sub_process_done_event.set()
                start_sub_process_event.clear()
                ProcessFunc.retries += 1
                continue

    def rad_button_click():
        procs = findwindows.find_elements()
        for proc_list in procs:
            if proc_list.control_type == "Telerik.WinControls.RadMessageBoxForm":
                for item in proc_list.children():
                    # if item.name == "입력한 고객과 동일한 고객 정보(이름, 휴대폰번호)가 존재합니다":
                    #     return
                    # elif item.name == "이름을 입력해주세요":
                    #     return
                    # elif item.name == "저장이 완료되었습니다":
                    if item.name == "확인":
                        btn = HwndWrapper(item)
                        btn.click()
                        break
