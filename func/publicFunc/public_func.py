from PIL import ImageGrab
from functools import partial
import os
import ctypes

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
