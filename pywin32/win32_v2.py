# coding: utf-8
# __author__: ""
from __future__ import unicode_literals

import win32gui
import win32con
import pyperclip
import time
import mouse_keyboard as mk
from pymouse import PyMouse


def open_child_wind(left, top, sender, sleep=0.7):
    '''
    打开子窗口
    :param left:
    :param top:
    :param sender:
    :param sleep:
    :return:
    '''
    # 进行用户搜索
    print("sender====={0}，窗口等待时间{1}".format(sender, sleep))
    m = PyMouse()
    m.click(left+140, top+40, 1)
    pyperclip.copy(sender)
    mk.press_hold_release(0.0001, "ctrl", "v")
    time.sleep(sleep)

    # 打开子聊天窗口
    m.click(left + 140, top + 120, 1)
    time.sleep(0.05)

    # 打开的如果是搜素窗口 则关闭该窗口， 重新加时打开
    hwnd = win32gui.FindWindow("FTSMsgSearchWnd", "微信")
    if hwnd > 0 and sleep < 1:
        print("close FTSMsgSearchWnd in {}".format(sleep))
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        open_child_wind(left, top, sender, sleep=sleep + 0.1)


def get_chat_window(sender, index=0):
    '''
    获取 窗口
    :param sender:
    :param index: 尝试长度
    :return:
    '''
    # 获取子窗口
    hwnd = win32gui.FindWindow("ChatWnd", sender)
    if hwnd > 0:
        return hwnd

    print("sender ======{0} 次 {1}".format(index, sender))
    if index > 3:
        raise Exception("打开子窗口失败")

    hwnd = win32gui.FindWindow("WeChatMainWndForPC", "微信")
    if hwnd < 0:
        raise Exception("程序未启动")

    # 获取主窗口
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(hwnd)
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)

    open_child_wind(left, top, sender)
    time.sleep(0.1)
    next_index = index + 1
    get_chat_window(sender, next_index)


def send_message(hwnd, message):
    """
    发送消息
    :param hwnd:
    :param message:
    :return:
    """
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception as e:
        print(e.message)

    pyperclip.copy(message)
    time.sleep(0.1)
    mk.press_hold_release(0.005, "ctrl", "v")
    time.sleep(0.005)
    mk.press_hold_release(0.001, "enter")


if __name__ == "__main__":
    t1 = time.time()
    senders = ["文件传输助手", "test1", "test2", "test3", "test4", "test5"][0:6]
    for sender in senders:
        hwnd = get_chat_window(sender)
        message = "测试微信\n\r请勿回复  编号 06-4: {}".format(sender)
        send_message(hwnd, message)

    print(time.time() - t1)
