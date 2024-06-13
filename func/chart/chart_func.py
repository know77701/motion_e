import random
from pywinauto import Application, findwindows
import time


class ChartFunc():
    """
        환자 차트 진입 후 동작
    """
    link_list = []
    save_btn = None
    save_edit = None
    delete_btn = None
    
    def chart_starter():
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
        chart_window = ChartFunc.find_window()

        if chart_window is not None:
            app = Application(backend="uia").connect(
                handle=chart_window.handle)
            side_chart_window = app.window(handle=chart_window.handle).child_window(
                class_name="Chrome_WidgetWin_0", found_index=1)

            for side_memo_list in side_chart_window.children():
                for item_list in side_memo_list.children():
                    if item_list.element_info.control_type == "Hyperlink":
                        ChartFunc.link_list.append(item_list)
                    if item_list.element_info.control_type == "Button" and item_list.element_info.name == "저장":
                        ChartFunc.save_btn = item_list
                    if item_list.element_info.control_type == "Edit" and item_list.element_info.name == "메모 입력":
                        ChartFunc.save_edit = item_list
            return ChartFunc.link_list, ChartFunc.save_btn, ChartFunc.save_edit

    def memo_save(index_number):
        ChartFunc.link_list, ChartFunc.save_btn, ChartFunc.save_edit = ChartFunc.find_memo_field()
        if ChartFunc.link_list is not None and ChartFunc.save_btn is not None and ChartFunc.save_edit is not None:

            ChartFunc.link_list[index_number].click_input()
            texts = ["테스트1", "테스트2", "테스트3", "테스트4", "테스트5",
                     "테스트6", "테스트7", "테스트8", "테스트9", "테스트10"]
            ran_number = random.randint(1, 1000)

            if ChartFunc.save_edit.is_enabled():
                for i in range(ran_number):
                    ran_text = random.choice(texts)
                    ChartFunc.save_edit.set_text(ran_text)
                    ChartFunc.save_btn.click()
        else:
            print("차트 미실행상태")
    def memo_update(index_number):
        return
    
    def side_memo_save(index):
        ChartFunc.memo_save(0)

    def resr_rece_memo_save():
        return
    
    def call_memo_save():
        ChartFunc.memo_save(2)

    def side_memo_update():
        
        return

    def side_memo_delete():

        return

    def resr_rece_memo_updqte():
        return

    def resr_rece_memo_delete():
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
