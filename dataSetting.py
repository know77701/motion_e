from pywinauto import application, Desktop, keyboard, findwindows
import time
import ctypes
import sys
import multiprocessing
import random


MAX_RETRY = 3


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, ' '.join(sys.argv), None, 1)
    sys.exit()

win32_app = application.Application(backend='win32')
MotionApp = application.Application(backend='uia')


def consulting_popup_close(quit_event, motion_window):
    chart_window = MotionApp.window(auto_id=motion_window)
    save_popup = chart_window.child_window(
        auto_id="RadMessageBox")
    print(save_popup)
    save_popup_btn = save_popup.child_window(auto_id="radButton1")
    print(save_popup_btn)
    save_popup_btn.click()
    time.sleep(1)
    quit_event.set()


def popup_btn_click():
    save_pen_chart_window = chart_window.child_window(
        auto_id="PenChartAddFrom")
    for _ in range(3):
        save_popup = save_pen_chart_window.child_window(
            auto_id="RadMessageBox")
        save_popup_btn = save_popup.child_window(auto_id="radButton1")
        save_popup_btn.click()
        time.sleep(1)


def pen_chart_process_func(quit_event):
    try:
        time.sleep(2)
        save_pen_chart_window = chart_window.child_window(
            auto_id="DrawingForm")
        pen_chart_image_template = save_pen_chart_window.child_window(
            auto_id='radListView')
        pen_chart_image_list = pen_chart_image_template.children()
        random_choice_image = random.choice(pen_chart_image_list)
        random_choice_image.click_input()
        time.sleep(1)
        save_btn = save_pen_chart_window.child_window(
            auto_id="radButtonSave")
        save_btn.click()

        time.sleep(2)
        save_pen_chart_detail_window = chart_window.child_window(
            auto_id="PenChartAddFrom")
        detail_save = save_pen_chart_detail_window.child_window(
            auto_id="buttonSave")
        process2 = multiprocessing.Process(
            target=popup_btn_click, args=())
        process2.start()
        detail_save.click()
        time.sleep(1)
        process2.terminate()
        process2.join()

        quit_event.set()
    except Exception as e:
        print(e)
        quit_event.set()


def med_picture_process_func(quit_event):
    time.sleep(2)
    picture_add_from = chart_window.child_window(
        auto_id="MedicalHistoryAddFrom")
    test_window = picture_add_from.child_window(auto_id="radPanel1")
    t = test_window.children()
    print(test_window)
    # change_image = picture_add_from.child_window(
    #     auto_id="buttonChangeImage")
    # print(change_image)
    # change_image.click_input()
    # time.sleep(2)
    # image_select_window = chart_window.child_window(title="열기")
    # test = image_select_window.children()

    # detail_save = picture_add_from.child_window(
    #     auto_id="buttonSave")
    # process2 = multiprocessing.Process(
    #     target=popup_btn_click, args=())
    # process2.start()
    # detail_save.click()
    # time.sleep(1)
    # process2.terminate()
    # process2.join()

    quit_event.set()


def process_func(window_title, button_auto_id):
    try:
        quit_event = multiprocessing.Event()
        MotionStarter.appConnect()
        while not quit_event.is_set():
            try:
                popup = MotionApp.window(
                    title=MotionStarter.VersionSearch(window_title))
                popup_button = popup.child_window(
                    auto_id=button_auto_id, control_type="Button")
                popup_button.wait('visible')
                popup_button.click()
                quit_event.set()
            except findwindows.ElementNotFoundError as e:
                print(f"버튼 미존재: {e}")
                pass
            except application.timings.TimeoutError as e:
                print(f"시간초과: {e}")
                pass
            time.sleep(1)
    except Exception as e:
        print("프로세스 실패", e)


def new_process(window_name, popup_name, btn_auto_id, click_btn_id):
    if __name__ == '__main__':
        quit_event = multiprocessing.Event()
        process1 = multiprocessing.Process(
            target=process_func, args=(popup_name, btn_auto_id, quit_event))
        process1.start()
        window_name.child_window(auto_id=click_btn_id).click()
        quit_event.set()
        process1.join()


class MotionStarter:
    def VersionSearch(value):
        windows = Desktop(backend="uia").windows()

        for window in windows:
            try:
                window_text = window.window_text()
                if value in window_text:
                    MotionTitle = window_text
                    return MotionTitle
            except Exception as e:
                print('버전 찾기 실패', e)

    @staticmethod
    def loginClick(title, id):
        login_window = win32_app.window(title=title)
        login_window.child_window(auto_id=id).click()

    @staticmethod
    def appTitleAction(title, btnName):
        win32_app.connect(path="C:\\Motion\\Motion_E\\Motion_E.exe")
        MotionStarter.loginClick(title, btnName)
        win32_app.kill()
        time.sleep(5)
        MotionApp.connect(path="C:\\Motion\\Motion_E\\Motion_E.exe")

    @staticmethod
    def appConnect(retries=0):
        try:
            if MotionStarter.VersionSearch('모션.ver'):
                MotionApp.connect(
                    path="C:\\Motion\\Motion_E\\Motion_E.exe")
                print('기존 앱 연결')
            elif MotionStarter.VersionSearch('로그인'):
                MotionStarter.appTitleAction('로그인', 'btnLogin')
                print('로그인 성공')
            else:
                win32_app.start("C:\\Motion\\Motion_E\\Motion_E.exe")
                time.sleep(2)
                MotionStarter.loginClick('로그인', 'btnLogin')
                time.sleep(3)
                MotionApp.connect(
                    path="C:\\Motion\\Motion_E\\Motion_E.exe")

        except application.ProcessNotFoundError as e:
            print("앱 찾기 실패 :", e)
            if retries < MAX_RETRY:
                retries += 1
                print(f"재시도 횟수: {retries}")
                win32_app.start("C:\\Motion\\Motion_E\\Motion_E.exe")
                MotionStarter.appConnect(retries)
            else:
                print("최대 재시도 횟수에 도달했습니다. 프로그램을 종료합니다.")
        except application.AppStartError:
            print("앱 미설치 또는 앱 미존재")


def main():
    MotionStarter.appConnect()
    motion_window = MotionApp.window(
        title=MotionStarter.VersionSearch('모션.ver'))
    index_number = 0
    while index_number <= MAX_RETRY:
        try:
            acpt_list = motion_window.child_window(
            auto_id="acpt-list", control_type="List")

            print(acpt_list.children())
            list_items = acpt_list.children(control_type="ListItem")
            print(f"접수카드 개수: {len(list_items)}")
            for _ in range(list_items):
                if list_items:
                    item = list_items[list_items]

                    text_controls = item.children(control_type="Text")

                    if text_controls:
                        first_text_control = text_controls[1]
                        first_text_control.click_input()
                else:
                    print("접수카드가 존재하지 않습니다.")

                chart_window = MotionApp.window(auto_id="tBeautyChartForm")

                # 사이드 메모입력
                time.sleep(1)
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

                # # 콜메모 입력
                side_links[2].click_input()
                chart_window2 = MotionApp.window(auto_id="tBeautyChartForm")

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
                # 사이드 메모 종료
                index_number += 1
                
                # 상담 시작
                consulting = chart_window.child_window(auto_id="spnlCnst")
                cnst_user = [
                    '심수빈', '이건', '설형일', '남종호', '탁대훈', '전은일', '강기혁', '남궁상호',
                    '정재훈', '최규영', '김지연', '류소원', '노승범', '배윤민', '백세미', '오인우',
                    '정진용', '이은지', '이혜라', '장여령', '정근화', '정예지', '표해남'
                ]
                cnst_id = consulting.child_window(auto_id="cmbCnstId")
                cnst_value = cnst_id.children()
                ran_cnst_user = random.choice(cnst_user)
                cnst_value[0].set_text(ran_cnst_user)

                acpt_id = consulting.child_window(auto_id="cmbAcptCfrId")
                acpt_value = acpt_id.children()
                ran_acpt_user = random.choice(cnst_user)
                acpt_value[0].set_text(ran_acpt_user)


                # 메모 입력 시작
                lb_memo = consulting.child_window(auto_id="tableLayoutPanel3")
                memo_edit = lb_memo.children()
                memo_edit[0].set_text('상담메모 테스트')
                memo_edit[1].set_text('어시메모 테스트')
                # 메모 입력 종료

                # 시술 선택-------------- list 추가 및 for문 작성해서 여러개? 패키지 하나?
                mopr_list = ['[여드름/색소] 여드름', '[스킨케어] 스킨케어']
                mopr = consulting.child_window(auto_id="txtSrchMopr")
                mopr_inupt = mopr.children()
                random_mopr = random.choice(mopr_list)
                mopr_inupt[0].set_text(random_mopr)
                keyboard.send_keys('{ENTER}')
                mpor_search = consulting.child_window(auto_id="gvRegMopr")
                mopr_list = mpor_search.children()
                mopr_list_choice = random.choice(mopr_list)
                mopr_list_choice.click_input()
                # 시술 선택 종료 -------------


                # 패키지 선택 ---------------------------
                pckg_list = ['리프팅 패키지']
                pckg = consulting.child_window(auto_id="txtSrchPckg")
                pckg_inupt = pckg.children()
                random_pckg = random.choice(pckg_list)
                pckg_inupt[0].set_text("")
                time.sleep(1)
                pckg_inupt[0].set_text(random_pckg)
                keyboard.send_keys('{ENTER}')

                pckg_list_window = consulting.child_window(auto_id="gvSrchPckgList")
                pckg_list_input = pckg_list_window.children()
                random_list_choice = random.choice(pckg_list_input)
                random_list_choice.click_input()
                add_pckg = consulting.child_window(
                    auto_id="btnAddMopr", control_type="Button")
                add_pckg.click()
                # 패키지 선택 종료 ----------------------------

                save_btn = consulting.child_window(
                    auto_id="btnSaveCnst", control_type="Button")
            break
        except findwindows.ElementNotFoundError :     
            document = motion_window.child_window(
            control_type="Document", found_index=0)
            document_children = document.children()

            try:
                buttons = []
                for doc_btn in document_children:
                    if doc_btn.element_info.control_type == "Button":
                        buttons.append(doc_btn)

                if len(buttons) >= 3:
                    third_button = buttons[2]
                    third_button.click()
                else:
                    print("버튼찾기 실패")
            except Exception as e:
                print("클릭 실패 :", e)




    # def test():
    #     quit_event = multiprocessing.Event()

    #     process1 = multiprocessing.Process(
    #         target=consulting_popup_close, args=(quit_event))
    #     process1.start()
    #     save_btn.click()
    #     process1.join()  # 새로운 프로세스가 종료될 때까지 기다립니다.


    # if __name__ == '__main__':
    #     test()
    # 상담 종료


    # 진료 시작
    # diag = chart_window.child_window(
    #     auto_id="spnlDiag")

    # # 담당의 지정 ------
    # doctor_list = ['(주)트라이업', '김지헌', '김다빈', '김보람', '김산호', '김한빛',
    #                '나혜은', '배석민', '변지혜', '김시별', '김승철', '유경화',
    #                '최영지', '이선희', '오승철'
    #                ]
    # drid = diag.child_window(auto_id='cmbChrgDrId')
    # drid_input = drid.children()
    # drid_random = random.choice(doctor_list)
    # drid_input[0].set_text(drid_random)
    # # 담당의 지정 종료 ----------

    # # 증상경과 입력
    # symt_memo_list = ['']
    # symt_memo_detail = diag.child_window(auto_id='pnlSymtPrgsDetail')
    # symt_memo_input = symt_memo_detail.children()
    # symt_memo_input[0].set_text("진료기록")
    # drid_memo_list = ['']
    # drid_memo_detail = diag.child_window(auto_id='pnlDiagMemoDetail')
    # drid_memo_input = drid_memo_detail.children()
    # drid_memo_input[0].set_text("증상경과")
    # # 증상경과 입력 종료 --------

    # # 처방 입력 > 처방 어떤걸로 해야할지 고민해야함.
    # search_prsc = diag.child_window(auto_id='txtSrchPrsc')
    # search_prsc_input = search_prsc.children()
    # search_prsc_input[0].set_text('감')
    # time.sleep(1)
    # prsc_list = diag.child_window(auto_id='gvRegPrsc')
    # prsc_list_input = prsc_list.children()
    # limit_list = prsc_list_input[:5]
    # random_prsc = random.choice(limit_list)
    # random_prsc.click_input()
    # # 처방 종료

    # diag_save_btn = diag.child_window(
    #     auto_id='btnSaveDiagWait', control_type="Button")
    # diag_save_btn.click()

    # # 진료 종료

    # 펜차트 시작
    # pen_chart = chart_window.child_window(
    #     auto_id="PenChartControl")
    # pen_chart_add_btn = pen_chart.child_window(auto_id='radButtonAdd')


    # def main():
    #     quit_event = multiprocessing.Event()

    #     process1 = multiprocessing.Process(
    #         target=pen_chart_process_func, args=(quit_event,))

    #     process1.start()
    #     pen_chart_add_btn.click()
    #     process1.terminate()
    #     process1.join()


    # if __name__ == '__main__':
    #     main()
    # 펜차트 종료

    # 진료사진 시작 > 추후 다시
    # medical_picture = chart_window.child_window(
    #     auto_id="MedicalPicturesControl")
    # medical_picture_add_btn = medical_picture.child_window(auto_id='radButtonAdd')


    # def med_picture_add_func():
    #     quit_event = multiprocessing.Event()

    #     process1 = multiprocessing.Process(
    #         target=med_picture_process_func, args=(quit_event,))

    #     process1.start()
    #     medical_picture_add_btn.click()
    #     process1.terminate()
    #     process1.join()


    # if __name__ == '__main__':
    #     med_picture_add_func()
    # 진료사진 종료

    # 수납 시작
