import random
from pywinauto import Application, findwindows, keyboard
import time
from pywinauto.controls.hwndwrapper import HwndWrapper
import datetime


class ChartFunc():
    """
        환자 차트 진입 후 동작
    """
    memo_link_list = []
    memo_save_btn = None
    memo_save_edit = None
    memo_delete_btn_list = []
    memo_content_list = []

    def chart_starter():
        # ChartFunc.side_memo_save()
        # ChartFunc.memo_update(0)
        ChartFunc.memo_delete(0)
        time.sleep(1)
    
        # ChartFunc.call_memo_save()
        # ChartFunc.call_memo_update(2)
        # ChartFunc.call_memo_delete(2)
        time.sleep(1)

    def window_resize(motion_app):
        return

    def find_window():
        procs = findwindows.find_elements()
        chart_window = None
        for window in procs:
            if window.control_type == "MotionChart.Chart_2.HMI.tBeautyChartForm":
                chart_window = window
        return chart_window

    def find_link():
        chart_window = ChartFunc.find_window()
        chart_hwnd_wrapper = HwndWrapper(chart_window.handle)
        chart_hwnd_wrapper.set_focus()

        if chart_window is not None:
            app = Application(backend="uia").connect(
                handle=chart_window.handle)
            side_chart_window = app.window(handle=chart_window.handle).child_window(
                class_name="Chrome_WidgetWin_0", found_index=1)

        for side_memo_list in side_chart_window.children():
            for item_list in side_memo_list.children():
                if item_list.element_info.control_type == "Hyperlink":
                    ChartFunc.memo_link_list.append(item_list)
                    
    def find_memo_field():
        now = datetime.datetime.now()
        current_year = now.strftime("%Y")
        chart_window = ChartFunc.find_window()

        if chart_window is not None:
            app = Application(backend="uia").connect(
                handle=chart_window.handle)
            side_chart_window = app.window(handle=chart_window.handle).child_window(
                class_name="Chrome_WidgetWin_0", found_index=1)

            for side_memo_list in side_chart_window.children():
                for item_list in side_memo_list.children():
                    if item_list.element_info.control_type == "Button" and item_list.element_info.name == "저장":
                        ChartFunc.memo_save_btn = item_list
                    if item_list.element_info.control_type == "Edit" and item_list.element_info.name == "메모 입력":
                        ChartFunc.memo_save_edit = item_list
                    if item_list.element_info.control_type == "Button" and item_list.element_info.name == "":
                        ChartFunc.memo_delete_btn_list.append(item_list)
                    if item_list.element_info.control_type == "Text" and item_list.element_info.name != "삭제보기" and not current_year in item_list.element_info.name:
                        ChartFunc.memo_content_list.append(item_list)

    def memo_save(index_number):
        try:
            ChartFunc.find_memo_field()
            if ChartFunc.memo_save_btn is not None and ChartFunc.memo_save_edit is not None:
                ChartFunc.find_link()
                memo_link = ChartFunc.memo_link_list[index_number]
                
                memo_link.click_input()
                texts = ["테스트1", "테스트2", "테스트3", "테스트4", "테스트5",
                        "테스트6", "테스트7", "테스트8", "테스트9", "테스트10"]
                ran_number = random.randint(1, 5)
                ChartFunc.memo_save_edit.set_text("테스트")
                if ChartFunc.memo_save_edit.is_enabled():
                    for i in range(ran_number):
                        ran_text = random.choice(texts)
                        ChartFunc.memo_save_edit.set_text(ran_text)
                        ChartFunc.memo_save_btn.click()
                print("메모 저장 성공")
            else:
                print("차트 미실행상태")
        except Exception as e:
            print(f"메모 저장 실패 :{e}")

    def memo_update(index_number):
        try:
            print("메모 수정 시작")
            save_btn = []
            ChartFunc.find_memo_field()
            content_list = []
            if ChartFunc.memo_content_list is not None and ChartFunc.memo_link_list is not None:
                for i in range(len(ChartFunc.memo_content_list)):
                    if i % 2 == 0:
                        content_list.append(ChartFunc.memo_content_list[i])
            random_item = random.choice(content_list)
            random_item.click_input()
            keyboard.send_keys("^a")
            keyboard.send_keys("수정 테스트")
            chart_window = ChartFunc.find_window()

            if chart_window is not None:
                app = Application(backend="uia").connect(
                    handle=chart_window.handle)
                side_chart_window = app.window(handle=chart_window.handle).child_window(
                    class_name="Chrome_WidgetWin_0", found_index=1)

                for side_memo_list in side_chart_window.children():
                    for item_list in side_memo_list.children():
                        if item_list.element_info.name == "저장" and item_list.element_info.control_type == "Button":
                            save_btn.append(item_list)
            if not save_btn:
                edit_btn = save_btn[1]
                edit_btn.click_input()
                print("메모 수정 완료")
        except Exception as e:
            print(f"메모수정 실패 : {e}")
            



    def memo_delete(index_number):
        ChartFunc.find_memo_field()

        btn_list = []
        if ChartFunc.memo_delete_btn_list is not None:
            for i in range(len(ChartFunc.memo_delete_btn_list)):
                if i % 2 != 0:
                    btn_list.append(ChartFunc.memo_delete_btn_list[i])
        btn_list[0].click()
        # for i in range(btn_list):
        #     i.click()
            

    def side_memo_save():
        print("사이드메모 저장 시작")
        ChartFunc.memo_save(0)
        time.sleep(1)
        print("사이드메모 저장 종료")

    def call_memo_save():
        ChartFunc.memo_save(2)

    def side_memo_update():
        ChartFunc.memo_update(0)
        return

    def side_memo_delete():

        return

    def resr_rece_memo_updqte():
        ChartFunc.memo_update(2)
        return

    def call_memo_update():

        return

    def call_memo_delete():
        return

    def past_chart_view():
        return

    def past_resr_veiw():
        return

    def resr_update():
        return
