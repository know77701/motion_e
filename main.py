import multiprocessing
import ctypes
import sys
from func.process_func.process_func import *
from func.start.motion_starter import *
from func.dashboard.dashboard import *
from func.publicFunc.public_func import *


if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, ' '.join(sys.argv), None, 1)
    sys.exit()

if __name__ == "__main__":
    start_sub_process_event = multiprocessing.Event()
    sub_process_done_event = multiprocessing.Event()
    
    process_func = ProcessFunc()
    

    main_process = multiprocessing.Process(
        target=process_func.main_process_func, args=(start_sub_process_event, sub_process_done_event))
    sub_process = multiprocessing.Process(
        target=process_func.sub_process_func, args=(start_sub_process_event, sub_process_done_event))
    # chart_sub_process = multiprocessing.Process(
    #     target=process_func.chart_sub_process, args=(start_sub_process_event,sub_process_done_event))
    
    
    main_process.start()
    sub_process.start()
    
    # chart_sub_process.start()

    main_process.join()
    sub_process.terminate()
    # chart_sub_process.terminate()
