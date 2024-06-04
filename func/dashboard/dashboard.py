from pywinauto import keyboard
from func.start.motion_starter import *
import time
import random
from func.dto.dto import DashboardDto

retries = 0


class DashBoard():
    """
    Motion E 차트 대시보드 동작
    """
    def dashboard_start(dto: DashboardDto):
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
        """
        DashBoard.notice_create(dto.motion_window)
        DashBoard.notice_delete(dto.motion_window, dto.motion_app)

        DashBoard.user_save(dto)
        

    def notice_create(motion_window):
        try:
            motion_window.child_window(
                auto_id='notice-content', control_type='Edit').type_keys('TEST{ENTER}')
            print("공지등록 완료")
            time.sleep(3)
        except Exception as err:
            keyboard.send_keys('{F5}')
            print("공지등록 실패")

    def notice_delete(motion_window, motion_app):
        try:
            motion_window.child_window(
                title='닫기', control_type='Button', found_index=0).click()
            time.sleep(2)

            rad = motion_app.window(auto_id="RadMessageBox")
            radBtn = rad.child_window(
                auto_id="radButton1", control_type="Button")
            radBtn.click()
            print("공지사항 삭제 성공")
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
        DashBoard.popup_view(dto.motion_window, dto.search_name)
        registration_window = dto.motion_app.window(
            title=MotionStarter.version_search('고객등록'))
        print("text_edit_popup 1")
        window_list = registration_window.children()
        edit_list = []
        save_btn = None
        for item in window_list:
            if item.element_info.control_type == "Window" and item.element_info.name == "고객등록":
                for child in item.children():
                    if child.element_info.control_type == "Pane":
                        for child_list in child.children():
                            for value in child_list.children():
                                if value.element_info.control_type == "Edit":
                                    edit_list.append(value)
                                if value.element_info.control_type == "Button" and value.element_info.name == dto.btn_title:
                                    save_btn = value
        print("text_edit_popup 2")
        time.sleep(2)
        print("text_edit_popup 3")
        user_name = edit_list[19]
        user_name.set_text(dto.search_name)
        sec_mobile_edit3 = edit_list[11]
        fst_mobile_edit2 = edit_list[13]
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
                    dto.phone_number[1:4])
                sec_mobile_edit3.set_edit_text(
                    dto.phone_number[4:8])
       
        save_btn.click()
        dto.sub_process_done_event.wait()

    def save_receipt_popup(dto: DashboardDto):
        try:
            dto.btn_title = "저장+접수"
            DashBoard.text_edit_popup(dto)
            print("여긴")
            time.sleep(1)
            DashBoard.receipt(dto)

        except Exception as e:
            window_screen_shot("save_receipt_popup_fail.jpg")
            if MotionStarter.version_search('고객등록'):
                print("왜안돌지")
                registration_window = dto.motion_app.window(
                    title=MotionStarter.version_search('고객등록'))
                close_btn = registration_window.child_window(
                    title="닫기", control_type="Button")
                close_btn.click()
                print(e)
            elif MotionStarter.version_search('접수'):
                receipt = dto.motion_app.window(
                    title=MotionStarter.version_search('접수'))
                close_btn = receipt.child_window(
                    title="닫기", control_type="Button")
                close_btn.click()
                print(e)

    def save_reserve_popup(dto: DashboardDto):
        try:
            dto.btn_title = "저장+예약"
            DashBoard.text_edit_popup(dto)
            time.sleep(1)

        except:
            window_screen_shot("save_reserve_popup_fail.jpg")
            if MotionStarter.version_search('고객등록'):

                receipt_window = dto.motion_app.window(
                    title=MotionStarter.version_search('고객등록'))
                close_btn = receipt_window.child_window(
                    title="닫기", control_type="Button")
                close_btn.click()
            else:
                top_menu = dto.motion_app.child_window(
                    auto_id="pnTop")
                dashboard_menu = top_menu.child_window(title="Dashboard")
                dashboard_menu.click_input()

    def user_save(dto: DashboardDto):
        try:
            print(dto)
            dto.btn_title = "저장"
            DashBoard.text_edit_popup(dto)
            time.sleep(1)

        except Exception as e:
            print('저장 실패 : ', e)
            window_screen_shot("user_save_fail.jpg")
            if MotionStarter.version_search('고객등록'):
                registration_window = dto.motion_app.window(
                    title=MotionStarter.version_search('고객등록'))
                close_btn = registration_window.child_window(
                    title="닫기", control_type="Button")
                close_btn.click()
            keyboard.send_keys('{F5}')

    def receipt_check(motion_window, chart_number):
        motion_web_window = motion_window.child_window(
            class_name="Chrome_RenderWidgetHostHWND", control_type="Document")
        web_window = motion_web_window.children()
        doc_list = []
        for child in web_window:
            if child.element_info.control_type == 'Document':
                doc_list.append(child)

        list_wrapper = doc_list[3].children(control_type="List")

        for item in list_wrapper:
            child_elements = item.children()
            for list_item in child_elements:
                items = list_item.children()
                for items_child in items:
                    compare_number = items_child.element_info.name
                    if chart_number in compare_number:
                        print(f"접수 확인: {compare_number}")
                        items_child.click_input()
                        break

    def reserve_check(motion_window, chart_number):
        motion_web_window = motion_window.child_window(
            class_name="Chrome_RenderWidgetHostHWND", control_type="Document")
        web_window = motion_web_window.children()
        doc_list = []
        for child in web_window:
            if child.element_info.control_type == 'Document':
                doc_list.append(child)

        list_wrapper = doc_list[2].children(control_type="List")

        for item in list_wrapper:
            child_elements = item.children()
            for list_item in child_elements:
                items = list_item.children()
                for items_child in items:
                    compare_number = items_child.element_info.name
                    if chart_number in compare_number:
                        print(f"예약 확인: {compare_number}")
                        items_child.click_input()
                        break

    def user_card_cancel(motion_window, chart_number, index_number):
        """
        Args
            index_number - 리스트 하위 고정 값
            - 예약취소 = 2 / 접수취소 = 3
        """
        try:
            motion_web_window = motion_window.child_window(
                class_name="Chrome_RenderWidgetHostHWND", control_type="Document")
            web_window = motion_web_window.children()
            doc_list = []
            for child in web_window:
                if child.element_info.control_type == 'Document':
                    doc_list.append(child)

            list_wrapper = doc_list[index_number].children(control_type="List")
            found_chat_number = False

            for item in list_wrapper:
                child_elements = item.children()
                for list_item in child_elements:
                    items = list_item.children()
                    for items_child in items:
                        compare_number = items_child.element_info.name
                        if chart_number in compare_number:
                            found_chat_number = True
                            break
                    if found_chat_number:
                        for item in items:
                            if chart_number in compare_number:
                                if item.element_info.name == "닫기":
                                    item.click()
                                    break
        except Exception as e:
            window_screen_shot("cancle_fail.jpg")
            print(e)

    def popup_cancle_action(window_name, popup_text):
        """
        예약 취소 시 발생되는 팝업 동작
        """
        try:
            for wrapper in window_name:
                if wrapper.element_info.name == popup_text:
                    popup = wrapper.children()
                    for child in popup:
                        if child.element_info.control_type == 'Group':
                            fr_child = child.children()
                            for child in fr_child:
                                if child.element_info.name == "예" and child.element_info.control_type == 'Button':
                                    child.click()
                                    break
        except Exception as e:
            print(e)

    def receipt_cancel(motion_window, chart_number):
        try:
            DashBoard.user_card_cancel(
                motion_window, chart_number, 3)
            motion_web_window = motion_window.child_window(
                class_name="Chrome_RenderWidgetHostHWND", control_type="Document")
            motion_web_window.wait(
                wait_for='exists enabled', timeout=30)
            DashBoard.popup_cancle_action(
                web_window, "접수를 취소 하시겠습니까? (접수취소는 예약데이터가 없을경우 접수정보가 삭제됩니다)")
            web_window = DashBoard.motion_web_window.children()
        except TimeoutError as e:
            print("타임 아웃 : ", e)
            return

    def reserve_cancel(motion_window, chart_number):
        try:
            DashBoard.user_card_cancel(motion_window, chart_number, 2)
            motion_web_window = motion_window.child_window(
                class_name="Chrome_RenderWidgetHostHWND", control_type="Document")
            web_window = motion_web_window.children()
            for child in web_window:
                if child.element_info.name == '저장' and child.element_info.control_type == 'Button':
                    child.click()
                    break
            cancel_popup = motion_web_window.children()
            DashBoard.popup_cancle_action(cancel_popup, "예약을 취소 하시겠습니까?")

        except TimeoutError as e:
            print("타임 아웃 : ", e)
            return

    def search_btn_click(motion_window, chat_number, btn_title):
        """
            btn_title (string): 예약하기 / 접수하기 텍스트 입력
        """
        DashBoard.search_user(motion_window, chat_number)
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
                        if chat_number in item.element_info.name:
                            for item_value in items:
                                # btn_title = 예약하기 / 접수하기
                                if item_value.element_info.name == btn_title and item_value.element_info.control_type == "Button":
                                    item_value.click()
                                    break

    def reserve(motion_window, chat_number, btn_title):
        DashBoard.search_btn_click(motion_window, chat_number, btn_title)
        motion_web_window = motion_window.child_window(
            class_name="Chrome_RenderWidgetHostHWND", control_type="Document")
        parent_child = motion_web_window.children()
        document_list = []
        for child in parent_child:
            if child.element_info.control_type == "Document":
                document_list.append(child)
        document_new_list = document_list[2]
        child_list = document_new_list.children()
        fr_combo = []
        memo_list = []
        btn_list = []
        for list in child_list:
            if list.element_info.control_type == "ComboBox":
                fr_combo.append(list)
            if list.element_info.control_type == "Edit":
                memo_list.append(list)
            if list.element_info.control_type == "Button":
                btn_list.append(list)

        fr_combo[0].click_input()
        fr_combo_items = fr_combo[0].children()
        filter_list = []
        for combo_items in fr_combo_items:
            item_children = combo_items.children()
            random_item = random.choice(item_children)
            print(random_item)
            for item in item_children:
                name = item.element_info.name
                if name == random_item.element_info.name:
                    item.select()
                    break
                # item.set_focus()

        # if filter_list:
        #     random_item = random.choice(filter_list)
        #     random_item.set_focus()
        #     random_item.click_input()
        # print(random_item)
        # for i in item_children:
        # item_children.click_input()
        # if random_item.element_info.name != "시간":
        # print("테스트")
        # i.click_input()

        # item = fr_combo_items.children()
        # print(item)

        # memo_list[4].set_text("예약메모 테스트")
        # memo_list[5].set_text("전달메모 테스트")
        # 예약버튼
        # btn_list[0].click_input()

    def receipt(dto: DashboardDto):
        time.sleep(3)
        receipt_window = dto.motion_app.window(
            title=MotionStarter.version_search('접수'))

        print("테스트")
        receipt_list = receipt_window.children()
        fr_list = receipt_list[0].children()
        receipt_btn = None
        for child in fr_list:
            if child.element_info.control_type == "Button" and child.element_info.name == "접수":
                receipt_btn = child

        print("테스트2")
        sec_list = receipt_list[1].children()
        edit_list = []
        for wrapper in sec_list:
            item_list = wrapper.children()
            for item in item_list:
                if item.element_info.control_type == "Edit":
                    edit_list.append(item)
        edit_list[0].set_text("직원메모 입력")
        edit_list[1].set_text("접수메모 입력")
        print(dto.start_sub_process_event)
        dto.start_sub_process_event.set()
        receipt_btn.click()
        dto.sub_process_done_event.wait()

        # reservation = cusm.child_window(auto_id="search-list")
# reservation.child_window(title = "예약하기",found_index = 0).click()

# time_check = cusm.child_window(auto_id="reservation")
# combo = time_check.child_window(control_type = "ComboBox",found_index = 0)
# combo.click_input()
# combo_list = combo.child_window(control_type = "List")
# post= combo_list.children()
# for kkee in post :
#     print(kkee)
#     ran_btn = random.choice(post)
#     if kkee.element_info_name != "시간":
#         kkee.click_input()

#     # if rrdk.element_info.name != "시간":
#     #     rrdk.click_input()
#     #     break