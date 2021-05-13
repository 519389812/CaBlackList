import pyautogui
import os
import tkinter.filedialog
import re
import tkinter


def yes_no_box(text: str, title: str) -> str:
    select = pyautogui.confirm(text=text, title=title, buttons=['是', '否'])
    if not select:
        os._exit(0)
    return select


def alert_box(text: str, title: str) -> None:
    pyautogui.alert(text=text, title=title, button='确认')


def date_box(text: str = '请输入日期，格式为YYYY-MM-DD', title: str = '日期', default: str = '') -> str:
    date = pyautogui.prompt(text=text, title=title, default=default)
    if not date:
        os._exit(0)
    return date


def ask_box(text: str = '', title: str = '', default: str = '') -> str:
    answer = pyautogui.prompt(text=text, title=title, default=default)
    if not answer:
        os._exit(0)
    return answer


def open_file_box(title='选择文件', initialdir='.', filetypes=None):
    window = tkinter.Tk()
    window.withdraw()  # 隐藏
    try:
        if filetypes is None:
            filetypes = [('所有文件', '*.*')]
        file_path = tkinter.filedialog.askopenfilename(title=title, initialdir=initialdir, filetypes=filetypes)
        if not file_path:
            os._exit(0)
        return file_path
    finally:
        window.destroy()  # 销毁


if __name__ == '__main__':
    pass
