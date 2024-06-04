from func.dto.dto import DashboardDto
from func.start.motion_starter import *
from func.dashboard.dashboard import *
from func.publicFunc.public_func import *

MAX_RETRY = 3
retries = 0

class ProcessFunc():
    rad_box = None
    retries = 0

    def main_process_func(start_sub_process_event, sub_process_done_event):

        win32_app = application.Application(backend='win32')
        motion_app = application.Application(backend='uia')
        MotionStarter.app_connect(win32_app, motion_app)
        motion_window = motion_app.window(
            title=MotionStarter.version_search('모션.ver'))

        dto = DashboardDto(motion_window, motion_app, "자동화QA4", "01074417631",
                           start_sub_process_event, sub_process_done_event, "", "필요값")
        
        # 서브프로세스 통신용
        dto.start_sub_process_event.set()
        
        # 서브프로세스 대기용
        dto.sub_process_done_event.wait()
        
        #여기서부터 시작
        


    def sub_process_func(start_sub_process_event, sub_process_done_event):
        start_sub_process_event.wait()
        win32_app = application.Application(backend='win32')
        motion_app = application.Application(backend='uia')
        MotionStarter.app_connect(win32_app, motion_app)
        motion_window = motion_app.window(
            title=MotionStarter.version_search('모션.ver'))
        sub_process_done_event.set()
        
        start_sub_process_event.clear()
    
        start_sub_process_event.wait()
        reg_window = motion_window.child_window(auto_id="FrmRegPatInfo", title="고객등록")
        while ProcessFunc.retries <= MAX_RETRY:
            try:
                reg_window.child_window
                for reg_wrapper in reg_window:
                    print(reg_wrapper)
                    reg_list = reg_wrapper.children()
                    for reg_items in reg_list:
                        print(reg_items)
                        if reg_items.element_info.control_type == "Pane":
                            for item in reg_items.children():
                                print(item)
                                if item.element_info.control_type == "Button" and item.element_info.name == "확인":
                                    print(item)
                                    item.click()
                                    break
                
                sub_process_done_event.set()
                start_sub_process_event.clear()

                start_sub_process_event.wait()
            except Exception as e:
                print("서브 모듈 동작 실패 : ", e)
                retries += 1
                continue
