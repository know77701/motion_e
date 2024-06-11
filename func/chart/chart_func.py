import random


class ChartFunc():
    """
        환자 차트 진입 후 동작
    """

    def side_memo_save(motion_window):
        chart_window = motion_window.window(auto_id="tBeautyChartForm")
        chart_window.wait(wait_for='exists enabled', timeout=30)
        chart_doc = chart_window.child_window(control_type="Document", found_index=1)
        chart_doc_child = chart_doc.children()
        side_edits = []
        side_buttons = []
        side_links = []
        for dco_elment in chart_doc_child:
            if dco_elment.element_info.control_type == "Edit":
                side_edits.append(dco_elment)
            if dco_elment.element_info.control_type == "Button":
                side_buttons.append(dco_elment)
            if dco_elment.element_info.control_type == "Hyperlink":
                side_links.append(dco_elment)
        side_edits[0].click_input()
        texts = ["테스트", "테스트2", "테스트3", "테스트4", "테스트5",
                "테스트6", "테스트7", "테스트8", "테스트9", "테스트10"]
        ran_number = random.randint(1, 10)

        if side_edits[0].is_enabled():
            for i in range(ran_number):
                ran_text = random.choice(texts)
                side_edits[0].set_text(ran_text)
                side_buttons[1].click()

    def resr_rece_memo_save():
        return

    def call_memo_save():
        
        return

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
