from pywinauto import keyboard
from func.start.motion_starter import *
import time
import random
from func.dto.dto import DashboardDto


class DashBoard():

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

        user_name = edit_list[19]
        user_name.set_text(dto.search_name)
        sec_mobile_edit3 = edit_list[11]
        fst_mobile_edit2 = edit_list[13]
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
        dto.start_sub_process_event.set()
        save_btn.click()
        dto.sub_process_done_event.wait()

    def save_receipt_popup(dto: DashboardDto):
        try:
            dto.btn_title = "저장+접수"
            DashBoard.text_edit_popup(DashboardDto)
            receipt_window = DashboardDto.motion_app.window(
                title=MotionStarter.version_search('접수'))

            edit_field = receipt_window.child_window(auto_id="radPanel6")
            receipt_btn = receipt_window.child_window(
                auto_id="btnAcpt", control_type="Button")
            receipt_memo = edit_field.child_window(control_type="Edit")
            user_memo = edit_field.child_window(control_type="Edit")

            receipt_memo.set_text("테스트")
            user_memo.set_text("테스트")

            dto.start_sub_process_event.set()
            receipt_btn.click()
            dto.sub_process_done_event.wait()

            web_window = DashboardDto.motion_window.child_winodw(
                title="Motion E web", control_type="Document")
            acpt_list = web_window.child_window(
                control_type="List", auto_id="acpt-list")
            list_childs = acpt_list.children()

        except Exception as e:
            window_screen_shot("save_receipt_popup_fail.jpg")
            if MotionStarter.version_search('고객등록'):
                registration_window = DashboardDto.motion_app.window(
                    title=MotionStarter.version_search('고객등록'))
                close_btn = registration_window.child_window(
                    title="닫기", control_type="Button")
                close_btn.click()
                print(e)
            elif MotionStarter.version_search('접수'):
                receipt = DashboardDto.motion_app.window(
                    title=MotionStarter.version_search('접수'))
                close_btn = receipt.child_window(
                    title="닫기", control_type="Button")
                close_btn.click()
                print(e)

    def save_reserve_popup(dto: DashboardDto):
        try:
            dto.btn_title = "저장+예약"
            DashBoard.text_edit_popup(DashboardDto)
            time.sleep(1)

        except:
            window_screen_shot("save_reserve_popup_fail.jpg")
            if MotionStarter.version_search('고객등록'):

                receipt_window = DashboardDto.motion_app.window(
                    title=MotionStarter.version_search('고객등록'))
                close_btn = receipt_window.child_window(
                    title="닫기", control_type="Button")
                close_btn.click()
            else:
                top_menu = DashboardDto.motion_app.child_window(
                    auto_id="pnTop")
                dashboard_menu = top_menu.child_window(title="Dashboard")
                dashboard_menu.click_input()

    def user_save(dto: DashboardDto):
        try:
            dto.btn_title = "저장"
            DashBoard.text_edit_popup(DashboardDto)
            time.sleep(1)

        except Exception as e:
            print('저장 실패 : ', e)
            window_screen_shot("user_save_fail.jpg")
            if MotionStarter.version_search('고객등록'):
                registration_window = DashboardDto.motion_app.window(
                    title=MotionStarter.version_search('고객등록'))
                close_btn = registration_window.child_window(
                    title="닫기", control_type="Button")
                close_btn.click()
            keyboard.send_keys('{F5}')

    def receipt_check(motion_window, chart_number):
        acpt_list = motion_window.child_window(
            auto_id="acpt-list", control_type="List")
        list_items = acpt_list.children(control_type="ListItem")
        for i in range(len(list_items)):
            item = list_items[i]
            child_elements = item.children()
            for child in child_elements:
                compare_number = child.element_info.name
                if compare_number == chart_number:
                    print(f"접수확인: {compare_number}")
                    chart_number.click_input()
                    break

    def reserve_check(motion_window, chart_number):
        rsrv_list = motion_window.child_window(
            auto_id="rsrv-list", control_type="List")
        list_items = rsrv_list.children(control_type="ListItem")
        for i in range(len(list_items)):
            item = list_items[i]
            child_elements = item.children()

            for child in child_elements:
                compare_number = child.element_info.name
                if compare_number == chart_number:
                    print(f"예약 확인: {compare_number}")
                    child.click_input()
                    break

    def user_card_cancel(motion_window, chart_number, auto_id):
        try:
            card_list = motion_window.child_window(
                auto_id=auto_id, control_type="List")
            card_list.wait(wait_for='exists enabled', timeout=20)

            list_item = card_list.children(control_type="ListItem")
            found_chat_number = False

            for item in list_item:
                child_elements = item.children()
                for child in child_elements:
                    compare_number = child.element_info.name
                    if compare_number == chart_number:
                        found_chat_number = True
                        break
                if found_chat_number:
                    for child in child_elements:
                        if compare_number == chart_number:
                            if child.element_info.name == "닫기":
                                child.click()
        except Exception as e:
            window_screen_shot("cancle_fail.jpg")
            print(e)

    def popup_cancle_action(window_name, popup_text):
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
                motion_window, chart_number, "acpt-list")
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
            DashBoard.user_card_cancel(
                motion_window, chart_number, "rsrv-list")
            motion_web_window = motion_window.child_window(
                class_name="Chrome_RenderWidgetHostHWND", control_type="Document")
            motion_web_window.wait(
                wait_for='exists enabled', timeout=30)
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
        # DashBoard.search_btn_click(motion_window, chat_number, btn_title)
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
        for list in child_list:
            if list.element_info.control_type == "ComboBox":
                fr_combo = list
                break
        fr_combo.click_input()
        fr_combo_items = fr_combo.children()
        for combo_items in fr_combo_items:
            item_children = combo_items.children()
            random_item = random.choice(item_children)
            print(random_item)
                
        # item = fr_combo_items.children()
        # print(item)

        # for wrapper in child_list:
        #     if wrapper.element_info.control_type == "List":
        #         wrapper_item = wrapper.children()
        #         for list_items in wrapper_item:
        #             items =  list_items.children()
        #             for item in items:
        #                 if item.element_info.name in chat_number:
        #                     for item_value in items:
        #                         if item_value.element_info.name == btn_title and item_value.element_info.control_type == "Button":
        #                             item_value.click()
        #                             break
