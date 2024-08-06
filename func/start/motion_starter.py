from pywinauto import application, Desktop
import time
from func.publicFunc.public_func import *

MAX_RETRY = 3


class MotionStarter():
    def __init__(self):
        self.backend = "uia"
        self.process_title = "Motion_E.exe"
        self.login_title = "로그인"
        self.btn_auto_id = "btnLogin"
        self.file_path = "C:\\Motion\\Motion_E\\Motion_E.exe"

    def version_search(self, search_title):
        windows = Desktop(backend=self.backend).windows()

        for window in windows:
            try:
                window_text = window.window_text()
                if search_title in window_text:
                    motion_title = window_text
                    return motion_title
            except Exception as e:
                print('버전 찾기 실패', e)
                window_screen_shot("version_search_fail")

    def login_click(self, win32_app, title, id):
        try:
            login_window = win32_app.window(title=title)
            login_window.child_window(auto_id=id).click()
        except Exception as e:
            print("로그인 클릭 실패")
            window_screen_shot("login_click_fail")

    def app_title_connect(self,win32_app, motion_app, title, btnName):
        try:
            win32_app.connect(path=self.process_title)
            self.login_click(win32_app, title, btnName)
            time.sleep(5)
            motion_app.connect(path=self.process_title)
        except Exception as e:
            print("타이틀 찾기 실패 : ", e)
            window_screen_shot("app_title_connect_fail")

    def app_connect(self,win32_app, motion_app, retries=0):
        try:
            if self.version_search('모션.ver'):
                motion_app.connect(path=self.process_title)
                print('기존 앱 연결')
            elif self.version_search(self.login_title):
                self.app_title_connect(
                    win32_app, motion_app, self.login_title, self.btn_auto_id)
                print('로그인 성공')
            else:
                win32_app.start(self.file_path)
                time.sleep(2)
                self.login_click(win32_app, self.login_title, self.btn_auto_id)
                time.sleep(3)
                motion_app.connect(path=self.process_title)

        except application.ProcessNotFoundError as e:
            print("앱 찾기 실패 :", e)
            window_screen_shot("app_connect_fail")
            if retries < MAX_RETRY:
                retries += 1
                print(f"재시도 횟수: {retries}")
                self.app_connect(self,win32_app, motion_app, retries)
            else:
                print("최대 재시도 횟수에 도달했습니다. 프로그램을 종료합니다.")
        except application.AppStartError:
            print("앱 미설치 또는 앱 미존재")
            window_screen_shot("app_connect_fail")
