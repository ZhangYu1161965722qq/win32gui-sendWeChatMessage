import win32gui, win32api, win32con,win32clipboard
# import win32com.client
import time

# # 获取窗口句柄
# def get_all_windowHandler(hwnd,map_hwnd_title):
#     if win32gui.IsWindow(hwnd) and win32gui.IsWindowEnabled(hwnd): #and win32gui.IsWindowVisible(hwnd):
#         map_hwnd_title[hwnd]=win32gui.GetWindowText(hwnd)

# # --函数：根据标题，获取窗口句柄、标题
# def funGetWindowHandler(filename_noExtension):
#     map_hwnd_title=dict()

#     # 列举出所有窗口的句柄
#     win32gui.EnumWindows(get_all_windowHandler,map_hwnd_title)

#     # print(map_hwnd_title)

#     list_handler=[]
#     for winHandler,winTitle in map_hwnd_title.items():
    
#         if winTitle.strip()[0:len(filename_noExtension)]==filename_noExtension:
#             list_handler.append(winHandler)
#     return list_handler

def setTextToClipboard(str_text):
    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()

    time.sleep(0.1)
    win32clipboard.SetClipboardText(str_text)

    win32clipboard.CloseClipboard()

def sendCtrlAndKey(winHandler,ascii_key):
    win32api.keybd_event(17, 0, 0, 0)   # 按下CTRL
    time.sleep(0.1)
    win32api.PostMessage(winHandler,win32con.WM_KEYDOWN,ascii_key)     # 按键
    time.sleep(0.1)
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)    # 放开CTRL

def main(concat,message):
    winHandler=win32gui.FindWindow('WeChatMainWndForPC','微信')
    win32gui.BringWindowToTop(winHandler)
    win32gui.ShowWindow(winHandler,win32con.SW_RESTORE) # 取消最小化
    # shell = win32com.client.Dispatch("WScript.Shell")
    # shell.SendKeys('%')

    win32gui.SetForegroundWindow(winHandler)    # 窗口激活

    # ctrl+F查找联系人
    sendCtrlAndKey(winHandler,70)

    setTextToClipboard(concat)

    win32gui.SetForegroundWindow(winHandler)    # 窗口激活

    # ctrl+v
    sendCtrlAndKey(winHandler,86)

    time.sleep(0.5)

    win32gui.PostMessage(winHandler, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)   #回车键
    win32gui.PostMessage(winHandler, win32con.WM_KEYUP, win32con.VK_RETURN, 0)

    setTextToClipboard(message)

    win32gui.SetForegroundWindow(winHandler)    # 窗口激活

    # ctrl+v
    sendCtrlAndKey(winHandler,86)

    time.sleep(0.1)

    win32gui.PostMessage(winHandler, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)   #回车键


if __name__=='__main__':
    main('文件传输助手','资源预警 一切正常')