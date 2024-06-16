from PIL import ImageGrab
from functools import partial
import os
import ctypes
from pywinauto import keyboard
import time
import pyautogui


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


def user_delete(start_sub_process_event, sub_process_done_event, motion_window):
    try:
        motion_window.set_focus()
        compare_window = motion_window.children()
        for windows in compare_window:
            for window_list in windows.children():
                if window_list.element_info.control_type == "MenuBar":
                    for list_item in window_list.children():
                        if list_item.element_info.control_type == "MenuItem" and list_item.element_info.name == "고객관리":
                            list_item.double_click_input()
                            pyautogui.click(46,114)
                            time.sleep(1)
        
        pane_list = None        
        top_field_list = None
        delete_btn = None
        radio_list = []
        for windows in compare_window:
            for window_list in windows.children():
                if window_list.element_info.control_type == "Pane":
                    pane_list = window_list
                    
        for pane_wrapper in pane_list.children():
            for pane in pane_wrapper.children():
                for items in pane.children():
                    for item in items.children():
                        for i in item.children():
                            if i.element_info.automation_id == "gvPat_List":
                                radio_list.append(i)
                            if i.element_info.automation_id == "radPanel2":
                                top_field_list = i.children()
        
        for list in top_field_list:
            for item in list.children():
                if item.element_info.control_type == "Edit" and item.element_info.automation_id == "txtSearch":
                    item.set_text("QA")
                    time.sleep(0.5)
                    keyboard.send_keys('{ENTER}')
                if item.element_info.control_type == "Button" and item.element_info.automation_id == "btnPat_Delete":
                    delete_btn = item
    
        for radio_btn in radio_list:
            for item_list in radio_btn.children():
                for item in item_list.children():
                    if item_list:
                        item.click_input()
                        break
                    
        
        start_sub_process_event.set()
        time.sleep(1)
        if delete_btn is not None:
            delete_btn.click()
        else :
            print("버튼 찾기 실패")
        sub_process_done_event.wait()
                                
    except Exception as e:
        print(e)
