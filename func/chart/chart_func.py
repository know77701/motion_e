import random
from pywinauto import Application, findwindows, keyboard
import time
from pywinauto.controls.hwndwrapper import HwndWrapper
import datetime
import pyautogui


class ChartFunc():
    """
        환자 차트 진입 후 동작
    """
    def __init__(self):
        self.memo_link_list = []
        self.memo_save_btn = None
        self.memo_save_edit = None
        self.memo_delete_btn_list = []
        self.memo_content_list = []
        self.chart_child_list = []
        self.side_memo_value = "사이드메모"
        self.call_memo_value = "콜메모"
        self.rsrv_memo_value = "예약/접수 메모"
        now = datetime.datetime.now()
        self.current_year = now.strftime("%Y")
        self.explore_child_list = []
        self.doctor_user = ["(주)트라이업", "김지헌", "강남언니", "김다빈", "김보람", "김산호", "김준하", "김한빛",
                    "나혜은", "배석민", "변지혜", "김시별", "심평원", "김승철", "오준일", "유경화", "최영지", "이선희", "오승철"]
        self.management_user = ["김미리", "김정수", "김태현", "최성우",
                        "이예진", "이준혁", "강은지", "정예령", "정은지", "신지윤"]
        self.consulting_user = ["심수빈", "이건", "설형일", "남종호", "김지헌", "탁재훈", "전은일", "강기혁", "남궁상호", "정재훈", "최규영",
                        "김의사", "김지연", "류소원", "노승범", "배윤민", "박세미", "오인우", "정진용", "이은지", "이혜라", "정근화", "정예지", "표해남"]
        self.nurse_user = ["설해남", "맹효뢰", "배석민", "송준", "정은용", "신효진", "이승현", "정세희", "지유진"]
        self.coordinator_user = ["류미희", "오혜선"]

    def chart_starter(self, start_sub_process_event, sub_process_done_event):
        # ChartFunc.side_memo_save()
        # ChartFunc.side_memo_update()
        # ChartFunc.side_memo_delete()

        # ChartFunc.call_memo_save(2)
        # ChartFunc.call_memo_update()
        # ChartFunc.call_memo_delete()
        # ChartFunc.past_resr_veiw(start_sub_process_event)
        # ChartFunc.rsrv_cancle(start_sub_process_event)
        self.penchart_template_save()
        

    def window_resize(self, motion_app):
        return

    def find_window(self):
        procs = findwindows.find_elements()
        chart_window = None
        for window in procs:
            if window.control_type == "MotionChart.Chart_2.HMI.tBeautyChartForm":
                chart_window = window
        return chart_window

    def return_window(self,index_number=None, class_name=None, auto_id=None):
        chart_list = None
        chart_window = self.find_window()
        app = Application(backend="uia").connect(
            handle=chart_window.handle)
        if class_name != None and index_number != None:
            chart_list = app.window(handle=chart_window.handle).child_window(
                class_name=class_name, found_index=index_number)
        elif auto_id != None and index_number != None:
            chart_list = app.window(handle=chart_window.handle).child_window(
                auto_id=auto_id, found_index=index_number)
        elif auto_id != None:
            chart_list = app.window(handle=chart_window.handle).child_window(
                auto_id=auto_id)
        return chart_list

    def find_link(self, index_number):
        chart_window = self.find_window()
        chart_hwnd_wrapper = HwndWrapper(chart_window.handle)
        chart_hwnd_wrapper.set_focus()

        if chart_window is not None:
            chart_list = self.return_window(
                index_number=index_number, class_name="Chrome_WidgetWin_0")
            for side_memo_list in chart_list.children():
                for item_list in side_memo_list.children():
                    if item_list.element_info.control_type == "Hyperlink":
                        self.memo_link_list.append(item_list)

    def find_field(self,index_number):
        if index_number == 0:
            batch_list = []
            chart_list = self.return_window(
                index_number=index_number, class_name="Chrome_WidgetWin_0")
            for chart_item in chart_list.children():
                if chart_item.element_info.control_type == "Document":
                    for item in chart_item.children():
                        if self.current_year in item.element_info.name:
                            self.chart_child_list.append(item)
                            index = chart_item.children().index(item)
                            if index + 2 < len(chart_item.children()):
                                next_item = chart_item.children()[index + 2]
                                self.chart_child_list.append(next_item)
                        if item.element_info.control_type == "Button":
                            self.chart_child_list.append(item)
            array_lenght = 4
            for i in range(0, len(self.chart_child_list), array_lenght):
                batch = self.chart_child_list[i:i + array_lenght]
                batch_list.append(batch)
            return batch_list

        if index_number == 1:
            side_chart_window = self.return_window(
                index_number=index_number, class_name="Chrome_WidgetWin_0")
            for side_memo_list in side_chart_window.children():
                for item_list in side_memo_list.children():
                    if item_list.element_info.control_type == "Button" and item_list.element_info.name == "저장":
                        self.memo_save_btn = item_list
                    if item_list.element_info.control_type == "Edit" and item_list.element_info.name == "메모 입력":
                        self.memo_save_edit = item_list
                    if item_list.element_info.control_type == "Button" and item_list.element_info.name == "":
                        self.memo_delete_btn_list.append(item_list)
                    if item_list.element_info.control_type == "Text" and item_list.element_info.name != "삭제보기" and not self.current_year in item_list.element_info.name:
                        self.memo_content_list.append(item_list)

    def memo_save(self, index_number, memo_value):
        try:
            print(f"{memo_value} 저장 시작")
            self.find_field(index_number)
            print(self.memo_save_btn)
            if self.memo_save_btn is not None and self.memo_save_edit is not None:
                self.find_link()
                memo_link = self.memo_link_list[index_number]

                memo_link.click_input()
                texts = ["테스트1", "테스트2", "테스트3", "테스트4", "테스트5",
                         "테스트6", "테스트7", "테스트8", "테스트9", "테스트10"]
                ran_number = random.randint(1, 100)
                self.memo_save_edit.set_text("테스트")
                if self.memo_save_edit.is_enabled():
                    for i in range(ran_number):
                        ran_text = random.choice(texts)
                        self.memo_save_edit.set_text(ran_text)
                        self.memo_save_btn.click()
                print(f"{memo_value} 저장 성공")
            else:
                print("차트 미실행상태")
        except Exception as e:
            print(f"{memo_value} 저장 실패 :{e}")

    def memo_update(self,memo_value):
        try:
            print(f"{memo_value} 수정 시작")
            save_btn = []
            self.find_field(1)
            content_list = []
            if self.memo_content_list is not None and self.memo_link_list is not None:
                for i in range(len(self.memo_content_list)):
                    if i % 2 == 0:
                        content_list.append(self.memo_content_list[i])
            random_item = random.choice(content_list)
            random_item.click_input()
            keyboard.send_keys("^a")
            keyboard.send_keys("수정 테스트")
            chart_window = self.find_window()

            if chart_window is not None:
                side_chart_window = self.return_window(
                    index_number=1, class_name="Chrome_WidgetWin_0")

                for side_memo_list in side_chart_window.children():
                    for item_list in side_memo_list.children():
                        if item_list.element_info.name == "저장" and item_list.element_info.control_type == "Button":
                            save_btn.append(item_list)
            if not save_btn:
                edit_btn = save_btn[1]
                edit_btn.click_input()
                print(f"{memo_value} 수정 완료")
        except Exception as e:
            print(f"{memo_value} 메모수정 실패 : {e}")

    def memo_delete(self, memo_value):
        try:
            print(f"{memo_value} 삭제 시작")
            self.find_field(1)

            btn_list = []
            if self.memo_delete_btn_list is not None:
                for i in range(len(self.memo_delete_btn_list)):
                    if i % 2 != 0:
                        btn_list.append(self.memo_delete_btn_list[i])

            random_btn = random.choice(btn_list)
            if random_btn is not []:
                random_btn.click()
            else:
                print("버튼미존재")

            chart_window = self.find_window()
            app = Application(backend="uia").connect(
                handle=chart_window.handle)
            side_chart_window = app.window(handle=chart_window.handle).child_window(
                class_name="Chrome_WidgetWin_0", found_index=1)

            if side_chart_window is not None:
                for side_memo_list in side_chart_window.children():
                    if side_memo_list.element_info.control_type == "Document":
                        for memo_item in side_memo_list.children():
                            for item_child in memo_item.children():
                                for i in item_child.children():
                                    if i.element_info.control_type == "Button" and i.element_info.name == "예":
                                        i.click()
                                        print(f"{memo_value} 삭제 완료")
        except Exception as e:
            print(f"메모 삭제 실패 : {e}")

    def side_memo_save(self):
        self.memo_save(0, self.side_memo_value)
        time.sleep(1)

    def side_memo_update(self):
        self.memo_update(self.side_memo_value)
        time.sleep(1)

    def side_memo_delete(self):
        self.memo_delete(self.side_memo_value)
        time.sleep(1)

    def resr_rece_memo_updqte(self):
        self.find_link()
        if self.memo_link_list is not None:
            memo_link = self.memo_link_list[1]
            memo_link.click_input()
        else:
            print(f"{self.rsrv_memo_value} 탭 이동 불가")
        self.memo_update(self.rsrv_memo_value)

    def call_memo_save(self):
        self.memo_save(2, self.call_memo_value)
        time.sleep(1)

    def call_memo_update(self):
        self.memo_update(self.call_memo_value)
        time.sleep(1)

    def call_memo_delete(self):
        self.memo_delete(self.call_memo_value)
        time.sleep(1)

    def past_chart_view(self):
        self.find_link(0)
        chart_link = self.memo_link_list[0]
        chart_link.click_input()

        chart_list = self.find_field(0)
        self.view_tab(935, 227)
        random_chart_list = random.choice(chart_list)
        chart_view_btn = None
        chart_time = None
        chart_min = None
        if random_chart_list is not None:
            for item in random_chart_list:
                if item.element_info.control_type == "Text" and self.current_year in item.element_info.name:
                    chart_time = item.element_info.name
                elif item.element_info.control_type == "Button":
                    chart_view_btn = item
                else:
                    chart_min = item.element_info.name
        chart_view_btn.click()
        window_list = self.return_window(
            index_number=0, class_name="Chrome_WidgetWin_0")
        chart_receipt_tiem = f"{chart_time} {chart_min}"

        for doc_list in window_list.children():
            if doc_list.element_info.control_type == "Document":
                for list_item in doc_list.children():
                    for item in list_item.children():
                        if item.element_info.control_type == "Group":
                            for i in item.children():
                                if i.element_info.control_type == "Button" and i.element_info.name == "예":
                                    i.click()
        receipt_window = self.return_window(auto_id="radScrollablePanel1")

        receipt_time = None
        for receipt_list in receipt_window.children():
            for list_item in receipt_list.children():
                for items in list_item.children():
                    for item in items.children():
                        if item.element_info.automation_id == "radPanel1":
                            for child in item.children():
                                if child.element_info.automation_id == "pnlDiagHide":
                                    for i in child.children():
                                        if i.element_info.control_type == "Pane":
                                            receipt_time = i.element_info.name
        parts = receipt_time.split()
        date_part = parts[0]
        time_part = parts[1]
        hour_min_str = parts[2]

        date_str = date_part.replace('-', '')
        hour_min_parts = hour_min_str.split(':')
        hour = int(hour_min_parts[0])
        if '오후' in time_part:
            hour += 12
        minute = int(hour_min_parts[1])
        compare_tiem = f"{date_str} {hour}{minute}"
        if compare_tiem == chart_receipt_tiem:
            print("과거 차트 진입 완료")
        else:
            print("차트진입 실패")

    def view_tab(x, y):
        pyautogui.moveTo(x=x, y=y)
        pyautogui.click()

    def explore_children(self,element, depth=0, max_depth=0, index_number=0):
        if self.explore_child_list is None:
            self.explore_child_list = []

        if depth > max_depth:
            return
        if depth == index_number:
            self.explore_child_list.append(element)
        # print(f"{depth} level = {element}")
        for child in element.children():
            self.explore_children(
                child, depth + 1, max_depth, index_number=index_number)

    # def public_code_get(,search_value):
    #     return

    # def get_user():
    #     return

    def random_value(array):
        return random.choice(array)

    def tab_view(self,tab_title):
        if tab_title == "예약":
            self.view_tab(761, 221)
        elif tab_title == "상담":
            self.view_tab(852, 221)
        elif tab_title == "진료":
            self.view_tab(937, 221)
        elif tab_title == "펜차트":
            self.view_tab(1027, 221)
        elif tab_title == "진료사진":
            self.view_tab(1117, 221)
        elif tab_title == "동의서":
            self.view_tab(1207, 221)
        elif tab_title == "수납":
            self.view_tab(1287, 221)

    def rsvr_update(self):
        self.rsrv_tab_view()
        resr_window = self.return_window(auto_id="예약")

        self.explore_children(
            resr_window, depth=0, max_depth=5, index_number=4)
        list_items = self.explore_child_list[2].children()

        change_btn = None

        rsrv_tiem_table = None
        rdrddiag_fld_cd_edit = None
        cmb_rsrv_mopr_tp_cd_edit = None
        cmb_rsrv_cfr_id_edit = None
        cmb_rsrv_chrg_dr_id = None
        cmb_diag_rn_dc = None
        edit_list = []

        rsrv_value = ["미지정", "아쿠아필/듀클레어", "제모", "피코슈어", "보톡스", "필러", "모피우스",
                      "울쎄라", "올리지오", "볼뉴머", "티타늄", "온다 리프팅", "인모드/슈링크", "F/U&진료"]
        rsrv_significant = ["미지정", "complain", "f/u", "상담"]
        rsrv_non_surgical = ["미지정", "필러", "윤곽/조각/지방파괴주사", "리프팅", "색소/홍조",
                             "여드름자국/여드름흉터", "여드름", "스킨케어", "스킨부스터", "체형", "일반진료", "다이어트진료", "제모", "염증주사"]
        ran_rsrv_value = self.random_value(rsrv_value)
        ran_rsrv_significant_value = self.random_value(rsrv_significant)
        ran_rsrv_non_surgical_value = self.random_value(rsrv_non_surgical)
        rna_doc_user_value = self.random_value(self.doctor_user)
        rna_consulting_user_value = self.random_value(
            self.consulting_user)

        for item in list_items:
            if item.element_info.control_type == "Pane" and item.element_info.automation_id == "panel30":
                rsrv_tiem_table = item
            if item.element_info.name == "예약 변경" and item.element_info.control_type == "Button":
                change_btn = item
            if item.element_info.control_type == "Edit":
                edit_list.append(item)
            if item.element_info.name == "RdrdDIAG_FLD_CD" and item.element_info.control_type == "ComboBox":
                rdrddiag_fld_cd_edit = item.children()
                rdrddiag_fld_cd_edit[0].set_text(ran_rsrv_value)
            if item.element_info.name == "cmbRsrvMoprTpCd" and item.element_info.control_type == "ComboBox":
                cmb_rsrv_mopr_tp_cd_edit = item.children()
                cmb_rsrv_mopr_tp_cd_edit[0].set_text(
                    ran_rsrv_non_surgical_value)
            if item.element_info.name == "cmbRsrvCfrId" and item.element_info.control_type == "ComboBox":
                cmb_rsrv_cfr_id_edit = item.children()
                cmb_rsrv_cfr_id_edit[0].set_text(rna_consulting_user_value)
            if item.element_info.name == "cmbRsrvChrgDrId" and item.element_info.control_type == "ComboBox":
                cmb_rsrv_chrg_dr_id = item.children()
                cmb_rsrv_chrg_dr_id[0].set_text(rna_doc_user_value)
            if item.element_info.name == "cmbDiag_RN_CD" and item.element_info.control_type == "ComboBox":
                cmb_diag_rn_dc = item.children()
                cmb_diag_rn_dc[0].set_text(ran_rsrv_significant_value)

        compare_time = f"{self.now.hour}:{self.now.minute}"
        self.explore_child_list = []
        self.explore_children(
            rsrv_tiem_table, depth=0, max_depth=5, index_number=4)
        if self.explore_child_list:
            while True:
                random_time_table = random.choice(self.explore_child_list)
                print(random_time_table)
                if compare_time > random_time_table.element_info.name:
                    continue
                # if random_time_table.element_info.name == "20:00" or random_time_table.element_info.name == "20:15" or random_time_table.element_info.name == "20:30" or random_time_table.element_info.name == "20:45" or random_time_table.element_info.name == "21:00" or random_time_table.element_info.name == "21:15" or random_time_table.element_info.name == "21:30":
                print(random_time_table.is_visible())
                if random_time_table.is_visible():
                    continue
                else:
                    random_time_table.click_input()
                    break

        # scroll_bar = None
        # for items in rsrv_tiem_table.children():
        #     for item in items.children():
        #         if item.element_info.control_type == "ScrollBar":
        #             scroll_bar = item
        # scroll_bar.wheel_mouse_input(wheel_dist=1)
        # scroll_bar.wheel_mouse_input(wheel_dist=-1)

        if len(edit_list) >= 2:
            fr_edit = edit_list[0]
            sec_edit = edit_list[1]

            time.sleep(0.5)
            fr_edit.set_text("전달메모 테스트")
            sec_edit.set_text("예약메모 테스트")
            change_btn.click()
        else:
            print("메모 입력 불가")

    def rsrv_cancle(self, start_sub_process_event):
        resr_window = self.return_window(auto_id="예약")

        self.explore_children(
            resr_window, depth=0, max_depth=5, index_number=4)
        sec_child_list = self.explore_child_list[1].children()
        list_items = self.explore_child_list[2].children()
        cancle_btn = None

        rsrv_list = []
        list_view_btn = None

        for sec_i in sec_child_list:
            for items in sec_i.children():
                if items.element_info.control_type == "Pane":
                    list_view_btn = items.children()

        list_view_btn[0].click()

        for sec_i in sec_child_list:
            for items in sec_i.children():
                if items.element_info.control_type == "Custom":
                    rsrv_list.append(items)

        random_rsrv_list = random.choice(rsrv_list)
        random_rsrv_list.click_input()

        for item in list_items:
            if item.element_info.name == "예약 취소" and item.element_info.control_type == "Button":
                cancle_btn = item
        start_sub_process_event.set()
        cancle_btn.click()
        # cancle_window = ChartFunc.return_window(auto_id="PopReservationCancel")

        # for list in cancle_window.children():
        #     print(list)

    def rsrv_save(self):
        return

    def receipt_save(self, start_sub_process_event):
    
        try:    
            receipt_window = self.return_window(auto_id="진료")
            self.explore_children(receipt_window, depth=0, max_depth=6, index_number=6)
            
            sec_list = self.explore_child_list[1]
            th_list = self.explore_child_list[2]
            five_list = self.explore_child_list[4]
            
            print_btn = None
            save_btn = None
            prescription_edit = None
            corporal_edit = None
            edit_list = []
            corporal_item_list = None
            
            
            for sec_item in sec_list.children():
                for items in sec_item.children():
                    if items.element_info.name == "처방전 출력" and items.element_info.control_type == "Button":
                        print_btn = items
                    if items.element_info.name == "저장" and items.element_info.control_type == "Button":
                        save_btn = items
        
                    
            # 처방 입력영역
            for th_item in th_list.children():
                for items in th_item.children():
                    if items.element_info.control_type == "Edit":
                        prescription_edit = items
            
            random_doc_user= random.choice(self.doctor_user)
            
            for five_item in five_list.children():
                for items in five_item.children():
                    if items.element_info.control_type == "Pane":
                        for item in items.children():
                            if item.element_info.control_type == "Edit":
                                edit_list.append(item)
                    # 상병입력
                    if items.element_info.control_type == "Edit" and items.element_info.automation_id == "txtSrchSick":
                        corporal_edit = items
                    # 담당의
                    if items.element_info.control_type == "ComboBox" and items.element_info.automation_id == "cmbChrgDrId":
                        sec_item = items.children()
                        sec_item[0].set_text(random_doc_user)
            
            if corporal_edit is not None:
                corporal_edit.set_text("감기")
                keyboard.send_keys("{ENTER}")
            else:
                print("상병입력 불가")
            
            if prescription_edit is not None:
                prescription_edit.set_text("프로")
            else:
                print("처방 입력 불가")
                
            if edit_list is not []:
                edit_list[0].set_text("증상경과 작성")
                edit_list[1].set_text("진료기록 작성")
            else:
                print("증상경과 진료기록 입력 불가")    
            panle_window = self.return_window(auto_id="spnlDiag")

            for window_list in panle_window.children():
                for list_item in window_list.children():
                    if list_item.element_info.control_type == "Pane" and list_item.element_info.automation_id == "pnlSetPrscOrdr":
                        for order_items in list_item.children():
                            if order_items.element_info.control_type == "Table":
                                corporal_item_list = random.choice(order_items.children())
                    if list_item.element_info.control_type == "Pane" and list_item.element_info.automation_id == "cpnlSetPrscSrchSick":
                        for items in list_item.children():
                            for item in items.children():
                                for item_child in item.children():
                                    for i in item_child.children():
                                        if i.element_info.control_type == "Table" and i.element_info.automation_id == "gvSrchSickList":
                                            prescription_item = random.choice(i.children())

            corporal_item_list.click_input()
            time.sleep(1)
            prescription_item.click_input()
            time.sleep(1)
            start_sub_process_event.set()
            save_btn.click()
        except findwindows.ElementNotFoundError:
            self.tab_view("진료")
            self.receipt_save(start_sub_process_event)
            
    def penchart_template_save(self):
        try:
            save_btn = None
            drawing_save_btn = None
            drawing_rotat_btn = None
            drawing_load_image_btn = None
            drawing_save_btn = None
            
            penchart_window = self.return_window(auto_id="펜차트")
            for penchart_list in penchart_window.children():
                for list_item in penchart_list.children():
                    if list_item.element_info.control_type == "Pane":
                        for items in list_item.children():
                            if items.element_info.control_type == "Button" and items.element_info.name == "등록":
                                save_btn = items
                                
            if save_btn is not None:
                save_btn.click()
            drawing_form = self.return_window(auto_id="DrawingForm")
            for form_list in drawing_form.children():
                print(form_list)
                if form_list.element_info.control_type == "Pane" and form_list.element_info.automation_id == "radPanel1":
                    for btn_item_list in form_list.children():
                        if btn_item_list.element_info.control_type == "Button" and form_list.element_info.automation_id == "radbuttonRotation":
                            drawing_rotat_btn = btn_item_list
                        if btn_item_list.element_info.control_type == "Button" and form_list.element_info.automation_id == "radButtonLoadImage":
                            drawing_load_image_btn = btn_item_list
                        if btn_item_list.element_info.control_type == "Button" and form_list.element_info.automation_id == "radButtonSave":
                            drawing_save_btn = btn_item_list
                if form_list.element_info.control_type == "Pane" and form_list.element_info.automation_id == "radPanel2":
                    for items in form_list.children():
                        for item in items.children():
                            
                            print(item)
        except findwindows.ElementNotFoundError as e:
            self.tab_view("펜차트")
            self.penchart_template_save()
            
        
    def penchart_get(self):
        return

    def penchart_update(self):
        return

    def penchart_delete(self):
        return