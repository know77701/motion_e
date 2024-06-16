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
    notice_value = '안내사항'

    def main_process_func(start_sub_process_event, sub_process_done_event):

        win32_app = application.Application(backend=ProcessFunc.win32_value)
        motion_app = application.Application(backend=ProcessFunc.uia_value)
        MotionStarter.app_connect(win32_app, motion_app)
        motion_window = motion_app.window(
            title=MotionStarter.version_search(ProcessFunc.motion_value))

        dto = DashboardDto(motion_window, motion_app, "QA9", "01074417631",
                           start_sub_process_event, sub_process_done_event, "", "")

        ProcessFunc.notice_popup_close(motion_app)

        # 서브프로세스 통신용
        dto.start_sub_process_event.set()

        # 서브프로세스 대기용
        dto.sub_process_done_event.wait()

        # DashBoard.dashboard_starter(dto)
        user_delete(start_sub_process_event, sub_process_done_event, motion_window)

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
                ProcessFunc.rad_button_click(
                    start_sub_process_event, sub_process_done_event)
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

    def rad_button_click(start_sub_process_event, sub_process_done_event):
        procs = findwindows.find_elements()
        for proc_list in procs:
            if proc_list.control_type == "Telerik.WinControls.RadMessageBoxForm":
                proc = proc_list.children()
                for item in proc:
                    item_tle = item.automation_id == "radLabel1"
                    item_btn = item.automation_id == "radButton1"
                    print(proc)
                    if item_tle == "삭제할 환자를 선택해주세요.":
                        print(proc)
                        
                    
                    # if item_tle == "radLabel1" and item.automation_id in "삭제할 환자를 선택해주세요.":
                    #     print("테스트")
                    #     if item.name == "확인" or item.name == "OK":
                    #         btn = HwndWrapper(item)
                    #         btn.click()
                    #         break
                    # if "삭제되었습니다." in item.name or "삭제할 환자를 선택해주세요." in item.name or "저장되었습니다" in item.name:
                    #     if item.name == "확인" or item.name == "OK":
                    #         btn = HwndWrapper(item)
                    #         btn.click()
                    #         break
                    # if "삭제에 대한 모든 책임은 병원에 있습니다. 삭제하시겠습니까?" in item.name:
                    #     if item.name == "확인" or item.name == "OK":
                    #         btn = HwndWrapper(item)
                    #         btn.click()
                    #         start_sub_process_event.set()
                    #         sub_process_done_event.wait()
                    #         break
                    # if "해당 고객을 삭제하시겠습니까?" in item.name:
                    #     if item.name == "확인" or item.name == "OK":
                    #         btn = HwndWrapper(item)
                    #         btn.click()
                    #         start_sub_process_event.set()
                    #         sub_process_done_event.wait()
                    #         break
                    
    def notice_popup_close(motion_app):
        procs = findwindows.find_elements()
        notice_procs = None
        for pro in procs:
            if pro.name == "안내사항" and pro.automation_id == "PopEventViewer":
                notice_procs = pro
                break
        
        if notice_procs is not None:
            notice_window = motion_app.window(title=MotionStarter.version_search(ProcessFunc.notice_value))
            for item_list in notice_window.children():
                print(item_list)
                if item_list.element_info.control_type == "TitleBar":
                    for item in item_list.children():
                        if item.element_info.control_type == "Button" and item.element_info.name == "닫기":
                            item.click()
                            break