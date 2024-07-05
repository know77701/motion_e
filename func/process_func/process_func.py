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
    sucess_value = False
    dto_search_name = "QA9"
    dto_phone_number = "01074417631"

    def main_process_func(start_sub_process_event, sub_process_done_event):

        win32_app = application.Application(backend=ProcessFunc.win32_value)
        motion_app = application.Application(backend=ProcessFunc.uia_value)
        MotionStarter.app_connect(win32_app, motion_app)
        motion_window = motion_app.window(
            title=MotionStarter.version_search(ProcessFunc.motion_value))

        dto = DashboardDto(motion_window, motion_app, ProcessFunc.dto_search_name, ProcessFunc.dto_phone_number,
                           start_sub_process_event, sub_process_done_event, "", "0123456789")

        # 서브프로세스 통신용
        dto.start_sub_process_event.set()

        # 서브프로세스 대기용
        dto.sub_process_done_event.wait()

        # ProcessFunc.notice_popup_close(motion_app)
        # DashBoard.dashboard_starter(dto)

        

        ChartFunc.chart_starter(dto.start_sub_process_event,
                                dto.sub_process_done_event)


    def sub_process_func(start_sub_process_event, sub_process_done_event):
        start_sub_process_event.wait()
        win32_app = application.Application(backend=ProcessFunc.win32_value)
        motion_app = application.Application(backend=ProcessFunc.uia_value)
        MotionStarter.app_connect(win32_app, motion_app)
        sub_process_done_event.set()

        # start_sub_process_event.clear()

        while ProcessFunc.retries <= ProcessFunc.MAX_RETRY:
            try:
                start_sub_process_event.wait()
                ProcessFunc.rad_button_click(
                    start_sub_process_event, sub_process_done_event)
                sub_process_done_event.set()
                # start_sub_process_event.clear()
                ProcessFunc.retries = 0
                continue

            except Exception as e:
                print("서브 모듈 동작 실패 : ", e)
                sub_process_done_event.set()
                # start_sub_process_event.clear()
                ProcessFunc.retries += 1
                continue

    def rad_button_click(start_sub_process_event, sub_process_done_event):
        item_tle = None
        item_btn = None
        while True:
            procs = findwindows.find_elements()
            for proc_list in procs:
                # if proc_list.automation_id == "":
                #     proc_cancle_window = proc_list.children()
                #     for i in proc_cancle_window:
                #         print(i)
                #     if proc_list.control_type == "Telerik.WinControls.RadMessageBoxForm":
                #         proc = proc_list.children()
                #         for item in proc:
                #             if item.automation_id == "radLabel1":
                #                 item_tle = item.name.strip().replace('\n', '').replace('\r', '').replace(' ', '')
                #             if item.automation_id == "radButton1":
                #                 item_btn = item
                # else:
                #     if proc_list.control_type == "Telerik.WinControls.RadMessageBoxForm":
                #         proc = proc_list.children()
                #         for item in proc:
                #             if item.automation_id == "radLabel1":
                #                 item_tle = item.name.strip().replace('\n', '').replace('\r', '').replace(' ', '')
                #             if item.automation_id == "radButton1":
                #                 item_btn = item
                return
            if item_tle is None:
                time.sleep(1)
                continue
            break

        if "삭제할환자를선택해주세요." in item_tle:
            item_btn = HwndWrapper(item_btn)
            item_btn.click()
        elif "이름을 입력하세요" in item_tle:
            item_btn = HwndWrapper(item_btn)
            item_btn.click()
        elif "삭제되었습니다." in item_tle or "저장되었습니다" in item_tle or "예약이완료되었습니다!" in item_tle or "접수완료되었습니다." in item_tle or "완료되었습니다." in item_tle or "예약되었습니다." in item_tle:
            item_btn = HwndWrapper(item_btn)
            item_btn.click()
            start_sub_process_event.clear()
        elif "접수하시겠습니까?" in item_tle:
            item_btn = HwndWrapper(item_btn)
            item_btn.click()
            start_sub_process_event.set()
            time.sleep(2)
            sub_process_done_event.wait()
        elif "예약을취소하시겠습니까?" in item_tle:
            item_btn = HwndWrapper(item_btn)
            item_btn.click()
            start_sub_process_event.set()
            time.sleep(2)
            sub_process_done_event.wait()
        elif "삭제에대한모든책임은병원에있습니다.삭제하시겠습니까?" in item_tle:
            item_btn = HwndWrapper(item_btn)
            item_btn.click()
            start_sub_process_event.set()
            time.sleep(2)
            sub_process_done_event.wait()
        elif "해당고객을삭제하시겠습니까?해당고객의처방및모든데이터가삭제됩니다." in item_tle:
            item_btn = HwndWrapper(item_btn)
            item_btn.click()
            start_sub_process_event.set()
            time.sleep(2)
            sub_process_done_event.wait()
        else:
            raise Exception('팝업 확인필요')

    def notice_popup_close(motion_app):
        procs = findwindows.find_elements()
        notice_procs = None
        for pro in procs:
            if pro.name == "안내사항" and pro.automation_id == "PopEventViewer":
                notice_procs = pro
                break

        if notice_procs is not None:
            notice_window = motion_app.window(
                title=MotionStarter.version_search(ProcessFunc.notice_value))
            notice_list = notice_window.children()
            for item_list in notice_list:
                if item_list.element_info.control_type == "TitleBar":
                    for item in item_list.children():
                        if item.element_info.control_type == "Button" and item.element_info.name == "닫기":
                            item.click()
                            break
