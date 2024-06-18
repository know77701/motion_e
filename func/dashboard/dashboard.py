from pywinauto import keyboard, findwindows
from func.start.motion_starter import *
import time
import random
from func.publicFunc.public_func import *
from func.dto.dto import DashboardDto
from func.chart.chart_func import *


class DashBoard():
    """
    Motion E 차트 대시보드 동작
    """
    RETRIES = 0
    MAX_RETRY = 3

    def dashboard_starter(dto: DashboardDto):
        """
        Args:
            dto (DashboardDto): 
            dto.motion_window = 모션 window 창
            dto.motion_app = process connect
            dto.search_name = 대시보드 검색 이름
            dto.phone_number = 환자 핸드폰 번호(고객 등록 시)
            dto.start_sub_process_event = 서브 프로세스 시작값
            dto.sub_process_done_event = 서브 프로세스 종료값
            dto.bnt_title = 신환 접수/예약/등록에 따라 변경되는 값
            dto.chart_number = 접수/예약 후 비교 숫자

        """

        # 화면 초기화
        DashBoard.dashboard_reset(dto.motion_window, dto.motion_app)

        # 공지사항 등록/비교/삭제
        DashBoard.notice_create(dto.motion_window)
        DashBoard.notice_delete(dto.motion_window, dto.motion_app)

        # 신환 등록
        DashBoard.user_save(dto)

        # 등록 환자 예약/비교
        dto.btn_title = "예약하기"
        DashBoard.search_btn_click(
            dto.motion_window, dto.chart_number, dto.btn_title)
        DashBoard.reserve(dto)
        DashBoard.reserve_cancel(dto.motion_window, dto.chart_number)


        # 등록 환자 접수/비교
        dto.btn_title = "접수하기"
        DashBoard.search_btn_click(
            dto.motion_window, dto.chart_number, dto.btn_title)
        DashBoard.receipt(dto)
        DashBoard.receipt_cancel(dto.motion_window, dto.chart_number)
        DashBoard.receipt_cancel(dto.motion_window, dto.chart_number)

        # # 고객등록 예약
        dto.search_name = dto.search_name + "예약"
        DashBoard.save_reserve_popup(dto)

        # 고객등록 접수
        dto.search_name = dto.search_name + "접수"
        DashBoard.save_receipt_popup(dto)
        
        # 환자차트 진입
        DashBoard.view_user_chart(dto.motion_window, 2, dto.chart_number)
        ChartFunc.chart_starter()

    def dashboard_reset(motion_window, motion_app):
        compare_window = motion_window.children()

        for window in compare_window:
            if window.element_info.name == "고객등록":
                for item in window.children():
                    if item.element_info.control_type == "TitleBar":
                        for title_bar in item.children():
                            if title_bar.element_info.name == "닫기" and title_bar.element_info.control_type == "Button":
                                title_bar.click()
                            if title_bar.element_info.name == "최대화" and title_bar.element_info.control_type == "Button":
                                title_bar.click()
                                break
        if MotionStarter.version_search('접수'):
            receipt_window = motion_app.window(
                title="접수", control_type="Window", auto_id="PopAcpt")
            for item in receipt_window.children():
                if item.element_info.control_type == "TitleBar":
                    for title_bar in item.children():
                        if title_bar.element_info.name == "닫기" and title_bar.element_info.control_type == "Button":
                            title_bar.click()
                            break
        for windows in compare_window:
            for window_list in windows.children():
                if window_list.element_info.automation_id == "menuBar":
                    for item in window_list.children():
                        if item.element_info.control_type == "MenuItem" and item.element_info.name == "Dashboard":
                            item.select()
                if window_list.element_info.automation_id == "TitleBar":
                    for item in window_list.children():
                        if window_list.elemen_info.automation_id == "Maximize-Restore" and item.element.control_type == "Button":
                            item.click()
                            break
        time.sleep(2)

    def notice_create(motion_window):
        try:
            notice_content = "테스트"
            notice_window = motion_window.child_window(
                auto_id='notice-content', control_type='Edit')
            notice_window.wait(wait_for='exists enabled', timeout=30)
            notice_window.type_keys(notice_content + "{ENTER}")

            motion_web_window = motion_window.child_window(
                class_name="Chrome_RenderWidgetHostHWND", control_type="Document")
            notice_list = motion_web_window.child_window(
                auto_id="notice-list", control_type="List")

            for list_item in notice_list.children():
                for item in list_item.children():
                    if item.element_info.control_type == "Text" and item.element_info.name == notice_content:
                        print("공지등록 완료")
            time.sleep(1)
        except Exception as err:
            keyboard.send_keys('{F5}')
            print("공지등록 실패", err)

    def notice_delete(motion_window, motion_app):
        try:
            motion_web_window = motion_window.child_window(
                class_name="Chrome_RenderWidgetHostHWND", control_type="Document")
            
            notice_list = motion_web_window.child_window(
                auto_id="notice-list", control_type="List")
            

            for list_item in notice_list.children():
                for item in list_item.children():
                    if item.element_info.control_type == "Text" and item.element_info.name == "TEST":
                        for item in list_item.children():
                            print(item)
                            if item.element_info.control_type == "Button" and item.element_info.name == "닫기":
                                item.click()                                               
                                rad = motion_app.window(auto_id="RadMessageBox")
                                radBtn = rad.child_window(
                                auto_id="radButton1", control_type="Button")
                                radBtn.click()
                                print("공지사항 삭제 성공")
                                break
            time.sleep(1)                    
        except Exception as err:
            print("공지사항 삭제 실패", err)

    @staticmethod
    def search_user(motion_window, search_name):
        motion_web_window = motion_window.child_window(
            class_name="Chrome_RenderWidgetHostHWND", control_type="Document")
        child_list = motion_web_window.children()
        search_btn = None
        for item in child_list:
            if item.element_info.control_type == "Edit":
                item.set_text("")
                item.set_text(search_name)
            if item.element_info.control_type == "Button" and item.element_info.name == "검색":
                search_btn = item
        search_btn.click()
        time.sleep(1)

    def popup_view(motion_window, search_name):
        DashBoard.search_user(motion_window, search_name)
        motion_web_window = motion_window.child_window(
            class_name="Chrome_RenderWidgetHostHWND", control_type="Document")
        child_list = motion_web_window.children()
        document_list = []
        for item in child_list:
            if item.element_info.control_type == "Document":
                document_list.append(item)
        document_item = document_list[0].children()
        for item in document_item:
            if item.element_info.control_type == "Button" and item.element_info.name == "환자 등록 후 예약":
                item.click()
                break

    def text_edit_popup(dto: DashboardDto):
        try:
            DashBoard.popup_view(dto.motion_window, dto.search_name)
            registration_window = dto.motion_app.window(
                title=MotionStarter.version_search('고객등록'))
            window_list = registration_window.children()
            edit_list = []
            save_btn = None
            chart_number = None
            for item in window_list:
                if item.element_info.control_type == "Window" and item.element_info.name == "고객등록":
                    for child in item.children():
                        if child.element_info.control_type == "Pane":
                            for child_list in child.children():
                                for value in child_list.children():
                                    if value.element_info.control_type == "Edit":
                                        edit_list.append(value)
                                        if value.element_info.automation_id == "txtChart_No":
                                            chart_number = value.element_info.name
                                    if value.element_info.control_type == "Button" and value.element_info.name == dto.btn_title:
                                        save_btn = value
            time.sleep(3)
            dto.chart_number = chart_number
            while DashBoard.RETRIES <= DashBoard.MAX_RETRY:
                if edit_list is not []:
                    user_name = edit_list[19]
                    sec_mobile_edit3 = edit_list[11]
                    fst_mobile_edit2 = edit_list[13]
                    DashBoard.RETRIES = 0
                    break
                else:
                    DashBoard.RETRIES += 1
                    print(DashBoard.RETRIES)

            dto.start_sub_process_event.set()
            match len(dto.phone_number):
                case 13:
                    fst_mobile_edit2.set_edit_text(
                        dto.phone_number[4:8])
                    sec_mobile_edit3.set_edit_text(
                        dto.phone_number[10:13])
                case 11:
                    fst_mobile_edit2.set_edit_text(
                        dto.phone_number[3:7])
                    sec_mobile_edit3.set_edit_text(
                        dto.phone_number[7:11])
                case 8:
                    fst_mobile_edit2.set_edit_text(
                        dto.phone_number[0:3])
                    sec_mobile_edit3.set_edit_text(
                        dto.phone_number[4:8])
            user_name.set_text(dto.search_name)
            save_btn.click()
            dto.sub_process_done_event.wait()

        except Exception as e:
            print(f"exception : {e}")
            window_screen_shot("text_edit_popup_fail")
            DashBoard.dashboard_reset(dto.motion_window, dto.motion_app)

    def save_receipt_popup(dto: DashboardDto):
        try:
            dto.btn_title = "저장+접수"
            DashBoard.text_edit_popup(dto)
            time.sleep(1)
            DashBoard.receipt(dto)
            time.sleep(1)

        except Exception as e:
            window_screen_shot("save_receipt_popup_fail")
            DashBoard.dashboard_reset(dto.motion_window, dto.motion_app)

    def save_reserve_popup(dto: DashboardDto):
        try:
            dto.btn_title = "저장+예약"
            DashBoard.text_edit_popup(dto)
            time.sleep(1)
            DashBoard.reserve(dto)
            time.sleep(1)

        except:
            window_screen_shot("save_reserve_popup_fail")
            DashBoard.dashboard_reset(dto.motion_window, dto.motion_app)

    def user_save(dto: DashboardDto):
        try:
            dto.btn_title = "저장"
            DashBoard.text_edit_popup(dto)
            time.sleep(2)
        except Exception as e:
            print('저장 실패 : ', e)
            window_screen_shot("user_save_fail")
            DashBoard.dashboard_reset(dto.motion_window, dto.motion_app)

    def return_list(motion_window,index_number):
        motion_web_window = motion_window.child_window(
            class_name="Chrome_RenderWidgetHostHWND", control_type="Document")
        web_window = motion_web_window.children()
        doc_list = []
        for child in web_window:
            if child.element_info.control_type == 'Document':
                doc_list.append(child)
        compare_item = doc_list[1]
        for item in compare_item.children():
            if item.element_info.name == "문자 발송":
                index_number += 1
                break
        
        doc_item = doc_list[index_number]
        list_wrapper = doc_item.children(control_type="List")
        return list_wrapper


    def card_check(motion_window, chart_number, index_number, card_type):
        # index_number 예약 = 1 / 접수 = 2
        try:
            list_wrapper = DashBoard.return_list(motion_window, index_number)
            for item in list_wrapper:
                child_elements = item.children()
                for list_item in child_elements:
                    items = list_item.children()
                    for items_child in items:
                        compare_number = items_child.element_info.name
                        if chart_number in compare_number:
                            print(f"확인: {compare_number}")
                            break
        except Exception as e:
            print(f"{card_type} 체크 실패 {e}")
            window_screen_shot("receipt_check_fail")
   
    def user_card_cancel(motion_window, chart_number, index_number):
        """
        Args
            index_number - 리스트 하위 고정 값
            - 예약취소 = 1 / 접수취소 = 2
        """
        try:
            list_wrapper = DashBoard.return_list(motion_window, index_number)
            found_chart_number = False

            for item in list_wrapper:
                child_elements = item.children()
                for list_item in child_elements:
                    items = list_item.children()
                    for items_child in items:
                        compare_number = items_child.element_info.name
                        if chart_number in compare_number:
                            found_chart_number = True
                            break
                    if found_chart_number:
                        for item in items:
                            if chart_number in compare_number:
                                if item.element_info.name == "닫기":
                                    item.click()
                                    break
        except Exception as e:
            window_screen_shot("cancle_fail")
            print(e)

    def popup_cancle_action(window_name, popup_text):
        """
            예약 취소 시 발생되는 팝업 동작
        """
        try:
            for wrapper in window_name:
                # if wrapper.element_info.name in popup_text:
                popup = wrapper.children()
                for child in popup:
                    if child.element_info.control_type == 'Group':
                        fr_child = child.children()
                        for i in fr_child:
                            if i.element_info.name == "예" and i.element_info.control_type == 'Button':
                                i.click()
                                break
        except Exception as e:
            print(e)
            window_screen_shot("popup_cancle_action_fail")

    def receipt_cancel(motion_window, chart_number):
        try:
            DashBoard.user_card_cancel(motion_window, chart_number, 2)
            motion_web_window = motion_window.child_window(
                class_name="Chrome_RenderWidgetHostHWND", control_type="Document")
            motion_web_window.wait(wait_for='exists enabled', timeout=30)
            cancel_popup = motion_web_window.children()
            DashBoard.popup_cancle_action(
                cancel_popup, "접수를 취소 하시겠습니까?")

        except TimeoutError as e:
            print("타임 아웃 : ", e)
            return
        except Exception as e:
            print(e)
            window_screen_shot("receipt_cancel")

    def reserve_cancel(motion_window, chart_number):
        try:
            DashBoard.user_card_cancel(motion_window, chart_number, 1)
            motion_web_window = motion_window.child_window(
                class_name="Chrome_RenderWidgetHostHWND", control_type="Document")
            motion_web_window.wait(wait_for='exists enabled', timeout=30)
            web_window = motion_web_window.children()
            for child in web_window:
                if child.element_info.name == '저장' and child.element_info.control_type == 'Button':
                    child.click()
                    break
            cancel_popup = motion_web_window.children()
            DashBoard.popup_cancle_action(cancel_popup, "예약을 취소 하시겠습니까?")

        except Exception as e:
            print(e)
            window_screen_shot("reserve_cancel")
            return

    def search_btn_click(motion_window, chart_number, btn_title):
        """
            btn_title (string): 예약하기 / 접수하기 텍스트 입력
        """
        DashBoard.search_user(motion_window, chart_number)
        motion_web_window = motion_window.child_window(
            class_name="Chrome_RenderWidgetHostHWND", control_type="Document")
        parent_child = motion_web_window.children()
        document_list = []
        for child in parent_child:
            if child.element_info.control_type == "Document":
                document_list.append(child)

        document_new_list = document_list[0]
        child_list = document_new_list.children()

        for wrapper in child_list:
            if wrapper.element_info.control_type == "List":
                wrapper_item = wrapper.children()
                for list_items in wrapper_item:
                    items = list_items.children()
                    for item in items:
                        if chart_number in items[0].element_info.name:
                            # btn_title = 예약하기 / 접수하기
                            if item.element_info.name == btn_title and item.element_info.control_type == "Button":
                                item.click()
                                break

        time.sleep(2)

    def reserve(dto: DashboardDto):
        try:
            time.sleep(1)
            motion_web_window = dto.motion_window.child_window(
                class_name="Chrome_RenderWidgetHostHWND", control_type="Document")
            parent_child = motion_web_window.children()
            document_list = []
            for child in parent_child:
                if child.element_info.control_type == "Document":
                    document_list.append(child)
            document_new_list = document_list[2]
            child_list = document_new_list.children()
            combo = []
            memo_list = []
            btn_list = []
            notice_list = []
            fr_key_press_counts = {
                "09": 1, "10": 2, "11": 3, "12": 4, "13": 5, "14": 6, "15": 7,
                "16": 8, "17": 9, "18": 10, "19": 11, "20": 12, "21": 13
            }
            sec_key_press_counts = {
                "00": 1, "15": 2, "30": 3, "45": 4
            }
            for list in child_list:
                if list.element_info.control_type == "ComboBox":
                    combo.append(list)
                if list.element_info.control_type == "Edit":
                    memo_list.append(list)
                if list.element_info.control_type == "Button":
                    btn_list.append(list)
                if list.element_info.control_type == "Document":
                    notice_list = list.children()

            combo[0].click_input()
            time.sleep(1.5)
            random_item = None
            while DashBoard.RETRIES <= DashBoard.MAX_RETRY:
                success = False
                for fr_item in combo[0].children():
                    if fr_item.is_visible():
                        item_children = fr_item.children()
                        random_item = random.choice(item_children)
                        if random_item.element_info.name == "시간":
                            continue
                        if random_item.element_info.name in fr_key_press_counts:
                            key_presses = fr_key_press_counts[random_item.element_info.name]
                            keyboard.send_keys('{DOWN}' * key_presses)
                            success = True
                            break
                if success:
                    break
                else:
                    DashBoard.RETRIES += 1
                    continue

            time.sleep(0.5)
            combo[1].click_input()
            while DashBoard.RETRIES <= DashBoard.MAX_RETRY:
                success = False
                for sec_item in combo[1].children():
                    if sec_item.is_visible():
                        sec_item_children = sec_item.children()
                        sec_random_item = random.choice(sec_item_children)
                        if random_item.element_info.name == "09":
                            if sec_random_item.element_info.name == "00":
                                continue
                        if sec_random_item.element_info.name == "분":
                            continue
                        if sec_random_item.element_info.name in sec_key_press_counts:
                            sec_key_presses = sec_key_press_counts[sec_random_item.element_info.name]
                            keyboard.send_keys('{DOWN}' * sec_key_presses)
                            success = True
                            break
                if success:
                    break
                else:
                    DashBoard.RETRIES += 1
                    continue

            time.sleep(1)
            if not memo_list[4].is_visible() or not memo_list[5].is_visible():
                if not notice_list == None:
                    notice_list[0].click_input()
            memo_list[4].set_text("예약메모 테스트")
            memo_list[5].set_text("전달메모 테스트")
            dto.start_sub_process_event.set()
            btn_list[0].click_input()

            dto.sub_process_done_event.wait()
            time.sleep(1)
            DashBoard.card_check(dto.motion_window, dto.chart_number, 1, "예약")
            time.sleep(1)
            compare_window = dto.motion_window.children()
            for windows in compare_window:
                for window_list in windows.children():
                    if window_list.element_info.automation_id == "menuBar":
                        for item in window_list.children():
                            if item.element_info.control_type == "MenuItem" and item.element_info.name == "Dashboard":
                                item.select()
                    if window_list.element_info.automation_id == "TitleBar":
                        for item in window_list.children():
                            if window_list.elemen_info.automation_id == "Maximize-Restore" and item.element.control_type == "Button":
                                item.click()
                                break
        except Exception as e:
            print(e)
            window_screen_shot("reserve_fail")

    def receipt(dto: DashboardDto):
        try:
            time.sleep(1)
            receipt_window = dto.motion_app.window(
                title="접수", control_type="Window", auto_id="PopAcpt")
            receipt_window.wait(wait_for='exists enabled', timeout=30)
            receipt_list = receipt_window.children()
            fr_list = receipt_list[0].children()
            sec_list = receipt_list[1].children()
            receipt_btn = None
            edit_list = []

            for item in fr_list:
                if item.element_info.control_type == "Button" and item.element_info.name == "접수":
                    receipt_btn = item
            for wrapper in sec_list:
                if wrapper.element_info.control_type == "Pane":
                    for item in wrapper.children():
                        if item.element_info.control_type == "Edit":
                            edit_list.append(item)

            time.sleep(0.5)
            dto.start_sub_process_event.set()
            edit_list[0].set_text("직원메모 입력")
            edit_list[1].set_text("접수메모 입력")
            receipt_btn.click()
            dto.sub_process_done_event.wait()
            time.sleep(0.5)
            DashBoard.card_check(dto.motion_window, dto.chart_number, 2, "접수")
        except Exception as e:
            print(e)
            window_screen_shot("receipt_fail")

    def view_user_chart(motion_window, index_number, chart_number):
        """
            index_number = 예약리스트 2 / 접수리스트 3
        """
        try:
            motion_web_window = motion_window.child_window(
                class_name="Chrome_RenderWidgetHostHWND", control_type="Document")
            motion_web_list = motion_web_window.children()
            chart_item = None
            doc_list = []
            for web_item in motion_web_list:
                if web_item.element_info.control_type == "Document":
                    doc_list.append(web_item)

            doc_item = doc_list[index_number].children()
            for items in doc_item:
                if items.element_info.control_type == "List":
                    for item in items.children():
                        for i in item.children():
                            if chart_number in i.element_info.name:
                                chart_item = i

            if not chart_item == None:
                chart_item.click_input()
                print(f"{chart_item.element_info.name} chart view")
            elif chart_number == 2:
                print(f"{chart_item.element_info.name} 예약카드 미존재")
            elif chart_number == 3:
                print(f"{chart_item.element_info.name} 접수카드 미존재")
        except Exception as e:
            print(e)
            window_screen_shot("view_user_chart_fail")
