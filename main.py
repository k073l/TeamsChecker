import sys
import os
import logging
import time

import winsound
import win32gui
import win32ui
from ctypes import windll
import pygetwindow as gw

import numpy as np
import cv2
from PIL import Image
import waiting


def screen_window(hwnd: int) -> str or None:
    """
    Original: https://stackoverflow.com/questions/19695214/python-screenshot-of-inactive-window-printwindow-win32gui
    Gets a screenshot of a window, saves it to assets/test{epoch}.png and if everything succeeded returns the name
    :param hwnd: Window handle as integer
    :return: Name of the screenshot or None, if it has failed
    """
    try:
        left, top, right, bot = win32gui.GetWindowRect(hwnd)
        w = right - left
        h = bot - top

        hwndDC = win32gui.GetWindowDC(hwnd)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()

        saveBitMap = win32ui.CreateBitmap()
        saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)

        saveDC.SelectObject(saveBitMap)

        result = windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 0)

        bmpinfo = saveBitMap.GetInfo()
        bmpstr = saveBitMap.GetBitmapBits(True)

        im = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1)

        win32gui.DeleteObject(saveBitMap.GetHandle())
        saveDC.DeleteDC()
        mfcDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)

        name = f"assets/test{int(time.time())}.png"
        if result == 1:
            im.save(name)
        return name
    except Exception as e:
        logging.warning(f"Exception occurred: {e}")
        return None


def find_image(im: np.ndarray, tpl: np.ndarray) -> tuple:
    """
    Original: https://stackoverflow.com/questions/29663764/determine-if-an-image-exists-within-a-larger-image-and-if-so-find-it-using-py
    :param im: Large image; operated upon
    :param tpl: Small image; template to look for in large image
    :return: x and y of found template in image or None, None, if not found
    """
    im = np.atleast_3d(im)
    tpl = np.atleast_3d(tpl)
    H, W, D = im.shape[:3]
    h, w = tpl.shape[:2]

    # --Numpy magic begins here--
    # I don't understand anything of it

    # Integral image and template sum per channel
    sat = im.cumsum(1).cumsum(0)
    tplsum = np.array([tpl[:, :, i].sum() for i in range(D)])

    # Calculate lookup table for all the possible windows
    iA, iB, iC, iD = sat[:-h, :-w], sat[:-h, w:], sat[h:, :-w], sat[h:, w:]
    lookup = iD - iB - iC + iA
    # Possible matches
    possible_match = np.where(np.logical_and.reduce([lookup[..., i] == tplsum[i] for i in range(D)]))

    # Find exact match
    for y, x in zip(*possible_match):
        if np.all(im[y + 1:y + h + 1, x + 1:x + w + 1] == tpl):
            return y + 1, x + 1

    return None, None


def is_len_small() -> bool:
    """
    Determines if user has opened meeting window or not
    :return: bool
    """
    return False if len(gw.getWindowsWithTitle('Microsoft Teams')) > 1 else True


if __name__ == '__main__':
    try:
        if len(sys.argv) > 1:
            if sys.argv[1] == '-d':
                logging.basicConfig(level=logging.DEBUG, format='%(levelname)s - %(asctime)s: %(message)s')
                logging.debug("Logging level set to DEBUG")
            else:
                logging.basicConfig(level=logging.INFO, format='%(message)s')
        else:
            logging.basicConfig(level=logging.INFO, format='%(message)s')

        if sys.platform != 'win32' and logging.getLogger().getEffectiveLevel() != 10:
            raise SystemExit("This will work only on Windows. If you think it is an error, try running the script "
                             "with -d")

        # Get windows' handles
        hwnds = gw.getWindowsWithTitle('Microsoft Teams')
        logging.debug(f"Teams windows detected: {len(hwnds)}")

        if len(hwnds) == 0:
            raise SystemExit("No Teams windows are present")
        for index, handle in enumerate(hwnds):
            logging.debug(f"Window {index}, title: {handle.title}")
            logging.debug(f"Window {index}, window handle: {handle._hWnd}")

        if logging.getLogger().getEffectiveLevel() != 10 or len(hwnds) == 1:
            main_hwnd = hwnds[0]._hWnd
        else:
            number = int(input(f"Pick window:\n {[index and handle.title for index, handle in enumerate(hwnds)]}"))
            main_hwnd = hwnds[number]._hWnd

        small_image = cv2.imread('assets/' + [image for image in os.listdir('assets') if image.startswith('join')][0])

        while True:
            waiting.wait(is_len_small)
            name = screen_window(main_hwnd)
            large_image = cv2.imread(name)
            x, y = find_image(large_image, small_image)
            logging.debug(f"Join button found on location: {x, y}")

            if logging.getLogger().getEffectiveLevel() != 10:
                os.remove(name)

            if (x and y) is not None:
                logging.debug("Found, alarming the user")
                winsound.Beep(520, 800)
                time.sleep(0.2)
                winsound.Beep(520, 800)
                time.sleep(0.2)
                winsound.Beep(520, 800)
                time.sleep(0.2)
    except KeyboardInterrupt:
        try:  # Cleanup, since sometimes a broken file was left, might help
            os.remove(name)
        except FileNotFoundError:
            pass
        raise SystemExit("Exiting..")
