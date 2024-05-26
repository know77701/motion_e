from pywinauto import application
from func.start.motion_starter import *
from func.dashboard.dashboard import *
from func.process_func.process_func import *
import ctypes
import sys
import multiprocessing


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
motion_app = application.Application(backend='uia')
MotionStarter.app_connect(win32_app, motion_app)

motion_window = motion_app.window(title=MotionStarter.version_search('모션.ver'))

if __name__ == "__main__":

    start_sub_process_event = multiprocessing.Event()
    sub_process_done_event = multiprocessing.Event()

    main_process = multiprocessing.Process(
        target=ProcessFunc.main_process_func, 
        args=(start_sub_process_event, 
              sub_process_done_event, 
              motion_app, 
              motion_window))
    sub_process = multiprocessing.Process(
        target=ProcessFunc.sub_process_func, 
        args=(motion_app, 
              start_sub_process_event, 
              sub_process_done_event, 
              "RadMessageBox", 
              "radButton1", 
              motion_app))

    main_process.start()
    sub_process.start()

    main_process.join()
    sub_process.terminate()
