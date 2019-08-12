# coding: utf-8
# __author__: ""
from __future__ import unicode_literals

import win32gui
import win32con
import win32api
import pyperclip
import time
import mouse_keyboard as mk
from pymouse import PyMouse


def open_child_wind(left, top, sender):
    '''
    打开子窗口
    :param left:
    :param top:
    :param sender:
    :return:
    '''
    # 进行用户搜索
    m = PyMouse()
    m.click(left+140, top+40, 1)

    pyperclip.copy(sender)
    mk.press_hold_release(0.0001, "ctrl", "v")
    time.sleep(0.4)
    # 打开子聊天窗口
    m.click(top_x + 140, top_y + 120, 1)
    m.click(top_x + 140, top_y + 120, 1)
    time.sleep(0.005)


def get_chart_window(sender):
    '''
    获取 窗口
    :param class_name:
    :param sender:
    :return:
    '''
    # 获取子窗口
    hwnd = win32gui.FindWindow("ChatWnd", sender)
    if hwnd > 0:
        return hwnd

    hwnd = win32gui.FindWindow("WeChatMainWndForPC", "微信")
    if hwnd < 0:
        raise Exception("程序未启动")

    # 获取主窗口
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(hwnd)
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)

    open_child_wind(left, top, sender)
    get_chart_window(sender)
    time.sleep(0.05)


def send_message(message, hwnd):
    """
    发送消息
    :param message:
    :param hwnd:
    :return:
    """
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    win32gui.SetForegroundWindow(hwnd)

    pyperclip.copy(message)
    time.sleep(0.1)
    mk.press_hold_release(0.005, "ctrl", "v")
    time.sleep(0.005)
    mk.press_hold_release(0.001, "enter")


def seed_message(message, senders, window_name="微信"):
    """
    发送微信消息
    :param message: 发送的内容
    :param senders: 发送者，即通过谁来发送
    :param window_name:
    :return:
    """
    # time.sleep(10)
    # 获取窗口， 设置窗口位置文件
    hwnd = win32gui.FindWindow("WeChatMainWndForPC", window_name)
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

    # x, y, w, h
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, top_x, top_y, window_w, window_h, win32con.SWP_SHOWWINDOW)
    win32gui.SetForegroundWindow(hwnd)




    # 进行用户搜索
    # mk.click("left", top_x+140, top_y+40, sleep=sleep)

    m = PyMouse()
    # time.sleep(0.005)
    m.click(top_x+140, top_y+40, 1)

    time.sleep(1)
    win32api.ClipCursor((0, 0, 10000, 1000))

    print(m.position())
    print(m.screen_size())

    pyperclip.copy(senders)
    mk.press_hold_release(sleep, "ctrl", "v")
    time.sleep(0.5)
    m.click(top_x+140, top_y+120, 1)
    # mk.click("left", top_x+140, top_y+120, sleep=sleep)
    #
    # # 送法消息

    pyperclip.copy(message)
    time.sleep(0.1)
    mk.press_hold_release(sleep, "ctrl", "v")
    time.sleep(0.005)
    mk.press_hold_release(sleep, "enter")
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 900, top_y, window_w, window_h, win32con.SWP_SHOWWINDOW)


t1 = time.time()
top_x, top_y, window_w, window_h = 900 +400, 100, 710, 500
seed_message("测试微信\n\r你好微信\n\r呵呵, 很好---------------测试最小化2", "文件")

print(time.time() - t1)
# x, y = win32api.GetCursorPos()
# print("mouse_x={0}, mouse_y={1}".format(x, y))
# win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
# print("left={0}, top={1}, right={2}, bottom={3}".format(left, top, right, bottom))

# time.sleep(0.005)
# win32api.SetCursorPos([top_x + 140, top_y + 120])
# time.sleep(0.001)
# win32api.mouse_event(win32con.MOUSEEVENT)
