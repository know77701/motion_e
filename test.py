from PIL import ImageGrab
from functools import partial

ImageGrab.grab = partial(ImageGrab.grab, all_screens=True)
img = ImageGrab.grab()

img.save("./fail/fail_screen.png")