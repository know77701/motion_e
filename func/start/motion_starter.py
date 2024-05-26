from pywinauto import application, Desktop
from func.start.motion_starter import MotionStarter
import time


MAX_RETRY = 3

class MotionStarter:
    def version_search(search_title):
        windows = Desktop(backend="uia").windows()

        for window in windows:
            try:
                window_text = window.window_text()
                if search_title in window_text:
                    motion_title = window_text
                    return motion_title
            except Exception as e:
                print('버전 찾기 실패', e)

    @staticmethod
    def login_click(win32_app, title, id):
        login_window = win32_app.window(title=title)
        login_window.child_window(auto_id=id).click()

    @staticmethod
    def app_title_connect(win32_app,motion_app,title, btnName):
        win32_app.connect(path="C:\\Motion\\Motion_E\\Motion_E.exe")
        MotionStarter.login_click(win32_app, title, btnName)
        win32_app.kill()
        time.sleep(5)
        motion_app.connect(path="C:\\Motion\\Motion_E\\Motion_E.exe")

    @staticmethod
    def app_connect(win32_app,motion_app,retries=0):
        try:
            if MotionStarter.version_search('모션.ver'):
                motion_app.connect(
                    path="C:\\Motion\\Motion_E\\Motion_E.exe")
                print('기존 앱 연결')
            elif MotionStarter.version_search('로그인'):
                MotionStarter.app_title_connect(win32_app, motion_app, '로그인', 'btnLogin')
                print('로그인 성공')
            else:
                win32_app.start("C:\\Motion\\Motion_E\\Motion_E.exe")
                time.sleep(1)
                MotionStarter.login_click('로그인', 'btnLogin')
                time.sleep(1)
                motion_app.connect(
                    path="C:\\Motion\\Motion_E\\Motion_E.exe")

        except application.ProcessNotFoundError as e:
            print("앱 찾기 실패 :", e)
            if retries < MAX_RETRY:
                retries += 1
                print(f"재시도 횟수: {retries}")
                win32_app.start("C:\\Motion\\Motion_E\\Motion_E.exe")
                MotionStarter.app_connect(win32_app, motion_app,retries)
            else:
                print("최대 재시도 횟수에 도달했습니다. 프로그램을 종료합니다.")
        except application.AppStartError:
            print("앱 미설치 또는 앱 미존재")
        except Exception as e:
            print("알수없는 에러 발생")
