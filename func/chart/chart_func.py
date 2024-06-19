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
    update_memo_content = "수정메모"

    def chart_starter():
        # ChartFunc.side_memo_save()
        ChartFunc.side_memo_save()

    def window_resize(motion_app):

        return

    def find_window():
        procs = findwindows.find_elements()
        chart_window = None
        for window in procs:
            if window.control_type == "MotionChart.Chart_2.HMI.tBeautyChartForm":
                chart_window = window
        return chart_window

    def find_memo_field():
        now = datetime.datetime.now()
        current_year = now.strftime("%Y")

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
                    if item_list.element_info.control_type == "Button" and item_list.element_info.name == "저장":
                        ChartFunc.memo_save_btn = item_list
                    if item_list.element_info.control_type == "Edit" and item_list.element_info.name == "메모 입력":
                        ChartFunc.memo_save_edit = item_list
                    if item_list.element_info.control_type == "Button" and item_list.element_info.name == "":
                        ChartFunc.memo_delete_btn_list.append(item_list)
                    if item_list.element_info.control_type == "Text" and item_list.element_info.name != "삭제보기" and not current_year in item_list.element_info.name:
                        ChartFunc.memo_content_list.append(item_list)

    def memo_save(index_number):
        ChartFunc.find_memo_field()
        if ChartFunc.memo_link_list is not None and ChartFunc.memo_save_btn is not None and ChartFunc.memo_save_edit is not None:
            ChartFunc.memo_link_list[index_number].click_input()
            time.sleep(2)

            texts = ["테스트1", "테스트2", "테스트3", "테스트4", "테스트5",
                     "테스트6", "테스트7", "테스트8", "테스트9", "테스트10"]
            ran_number = random.randint(1, 5)

            for i in range(ran_number):
                print(i)
                ran_text = random.choice(texts)
                print(ran_text)
                ChartFunc.memo_save_edit.set_text(ran_text)
                print(ChartFunc.memo_save_edit.element_info.name)
                ChartFunc.memo_save_btn.click()
                print(ChartFunc.memo_save_btn.element_info.name)
        else:
            print("차트 미실행상태")
        ChartFunc.memo_update()

    def memo_update():
        ChartFunc.find_memo_field()
        content_list = []
        if ChartFunc.memo_content_list is not None and ChartFunc.memo_link_list is not None:
            for i in range(len(ChartFunc.memo_content_list)):
                if i % 2 == 0:
                    content_list.append(ChartFunc.memo_content_list[i])
        random_item = random.choice(content_list)

        random_item.click_input()
        keyboard.send_keys('^a'+ChartFunc.update_memo_content)

        ChartFunc.find_memo_field()
        ChartFunc.memo_save_btn.click()
        ChartFunc.memo_compare(ChartFunc.update_memo_content)
        ChartFunc.memo_delete()

    def memo_delete():
        ChartFunc.find_memo_field()

        btn_list = []
        if ChartFunc.memo_link_list is not None and ChartFunc.memo_delete_btn_list is not None:
            for i in range(len(ChartFunc.memo_delete_btn_list)):
                if i % 2 != 0:
                    btn_list.append(ChartFunc.memo_delete_btn_list[i])
        ran_item = random.choice(btn_list)

        # for item in btn_list:
        #     if item == ran_item:
        ran_item.click_input()
        ChartFunc.delete_popup_action()

    def memo_compare(memo_content):
        ChartFunc.find_memo_field()
        if ChartFunc.memo_content_list is not None and ChartFunc.memo_link_list is not None:
            for content_item in ChartFunc.memo_content_list:
                if content_item.element_info.name == memo_content:
                    print(f"{memo_content} 메모확인")
                    break

    def delete_popup_action():
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
                    for item in item_list.children():
                        for i in item.children():
                            if i.element_info.name == "예" and i.element_info.control_type == "Button":
                                i.click()
                                print("사이드 메모 삭제 완료")
                                break

    def side_memo_save():
        ChartFunc.memo_save(0)

    def call_memo_save():
        ChartFunc.memo_save(2)

    def past_chart_view():
        return

    def past_resr_veiw():
        return

    def resr_update():
        return
