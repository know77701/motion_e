from pywinauto import keyboard
from func.start.motion_starter import *
import time


class DashBoard():

    serach_window = None
    register_btn = None
    search_btn = None
    register_btn = None
    registration_window = None
    edit_window = None
    fst_mobile_edit2 = None
    sec_mobile_edit3 = None

    @staticmethod
    def search_user(search_name, motion_window):
        DashBoard.serach_window = motion_window.child_window(
            auto_id="srch-val",  control_type="Edit")
        DashBoard.serach_window.set_edit_text("")
        DashBoard.serach_window.set_edit_text(search_name)
        DashBoard.search_btn = motion_window.child_window(
            title="검색", control_type="Button")
        DashBoard.search_btn.click()

    def popup_view(search_name, motion_window):
        DashBoard.search_user(search_name, motion_window)
        DashBoard.register_btn = motion_window.child_window(
            title="환자 등록 후 예약", control_type="Button")
        DashBoard.register_btn.wait(wait_for='exists enabled', timeout=30)
        DashBoard.register_btn.click()

    def text_edit_popup(serach_name, phone_number, start_sub_process_event, sub_process_done_event, btn_auto_id, motion_app, motion_window):
        DashBoard.popup_view(serach_name, motion_window)
        registration_window = motion_app.window(
            title=MotionStarter.version_search('고객등록'))
        edit_window = registration_window.child_window(
            control_type="Edit", auto_id="txtPat_Nm")
        edit_window.set_edit_text(serach_name)
        if phone_number:
            DashBoard.fst_mobile_edit2 = registration_window.child_window(
                control_type="Edit", auto_id="txtMobile_No2")
            DashBoard.sec_mobile_edit3 = registration_window.child_window(
                control_type="Edit", auto_id="txtMobile_No3")
            match len(phone_number):
                case 13:
                    DashBoard.fst_mobile_edit2.set_edit_text(phone_number[4:8])
                    DashBoard.sec_mobile_edit3.set_edit_text(
                        phone_number[10:13])
                case 11:
                    DashBoard.fst_mobile_edit2.set_edit_text(phone_number[3:7])
                    DashBoard.sec_mobile_edit3.set_edit_text(
                        phone_number[7:11])
                case 8:
                    DashBoard.fst_mobile_edit2.set_edit_text(phone_number[1:4])
                    DashBoard.sec_mobile_edit3.set_edit_text(phone_number[4:8])
        save_btn = registration_window.child_window(
            control_type="Button", auto_id=btn_auto_id)
        start_sub_process_event.set()
        save_btn.click()
        sub_process_done_event.wait()

    def save_receipt_popup(serach_name, phone_number, start_sub_process_event, sub_process_done_event, btn_auto_id, motion_app, motion_window):
        try:
            DashBoard.text_edit_popup(
                serach_name, phone_number, start_sub_process_event, sub_process_done_event, btn_auto_id, motion_app, motion_window)
            receipt_window = motion_app.window(
                title=MotionStarter.version_search('접수'))
            edit_field = receipt_window.child_window(auto_id="radPanel6")
            receipt_memo = edit_field.child_window(control_type="Edit")
            user_memo = edit_field.child_window(control_type="Edit")
            receipt_memo.set_text("테스트")
            user_memo.set_text("테스트")
            receipt_btn = receipt_window.child_window(
                auto_id="btnAcpt", control_type="Button")
            start_sub_process_event.set()
            receipt_btn.click()
            sub_process_done_event.wait()

            web_window = motion_app.child_winodw(title="Motion E web")
            acpt_list = web_window.child_window(
                control_type="List", auto_id="acpt-list")
            test = acpt_list.children()
            print(test)

        except Exception as e:
            if MotionStarter.version_search('고객등록'):
                registration_window = motion_app.window(
                    title=MotionStarter.version_search('고객등록'))
                close_btn = registration_window.child_window(
                    title="닫기", control_type="Button")
                close_btn.click()
                print(e)
            elif MotionStarter.version_search('접수'):
                receipt = motion_app.window(
                    title=MotionStarter.version_search('접수'))
                close_btn = receipt.child_window(
                    title="닫기", control_type="Button")
                close_btn.click()
                print(e)

    def save_reserve_popup(serach_name, phone_number, start_sub_process_event, sub_process_done_event, btn_auto_id, motion_app, motion_window):
        try:
            DashBoard.text_edit_popup(
                serach_name, phone_number, start_sub_process_event, sub_process_done_event, btn_auto_id, motion_app, motion_window)
            time.sleep(1)

        except:
            if MotionStarter.version_search('고객등록'):
                receipt_window = motion_app.window(
                    title=MotionStarter.version_search('고객등록'))
                close_btn = receipt_window.child_window(
                    title="닫기", control_type="Button")
                close_btn.click()
            else:
                top_menu = motion_app.child_window(auto_id="pnTop")
                dashboard_menu = top_menu.child_window(title="Dashboard")
                dashboard_menu.click_input()

    def user_save(serach_name, phone_number, start_sub_process_event, sub_process_done_event, btn_auto_id, motion_app, motion_window):
        try:
            DashBoard.text_edit_popup(
                serach_name, phone_number, start_sub_process_event, sub_process_done_event, btn_auto_id, motion_app, motion_window)
            time.sleep(1)

        except Exception as e:
            print('저장 실패 : ', e)
            if MotionStarter.version_search('고객등록'):
                registration_window = motion_app.window(
                    title=MotionStarter.version_search('고객등록'))
                close_btn = registration_window.child_window(
                    title="닫기", control_type="Button")
                close_btn.click()
            keyboard.send_keys('{F5}')

    def receipt_check(chart_number, motion_window):
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
                    break

    def reserve_check(chart_number, motion_window):
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
                    break
