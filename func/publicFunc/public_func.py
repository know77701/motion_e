from PIL import ImageGrab
from functools import partial
import os
import ctypes
from pywinauto import application, findwindows
import time

screenshot_save_dir = "fail"


def window_screen_shot(save_file_name):
    ImageGrab.grab = partial(ImageGrab.grab, all_screens=True)
    screenshot_path = os.path.join(
        screenshot_save_dir, save_file_name + ".jpg")
    save_image = ImageGrab.grab()
    save_image.save(screenshot_path)


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        window_screen_shot("login_click_fail")
        return False


def user_delete(motion_window):
    try:
        compare_window = motion_window.children()
        for windows in compare_window:
            for window_list in windows.children():
                if window_list.element_info.control_type == "MenuBar":
                    for list_item in window_list.children():
                        if list_item.element_info.control_type == "MenuItem" and list_item.element_info.name == "고객관리":
                            list_item.double_click_input()

        time.sleep(2)
        procs = findwindows.find_elements()

        for proc_list in procs:
            if proc_list.control_type == "Telerik.WinControls.UI.RadDropDownMenu":
                return
    except Exception as e:
        return
