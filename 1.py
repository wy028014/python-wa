import pyautogui
import pygetwindow as gw

window_name = "正则.txt - 记事本"

# 打印所有窗口标题，用于确认实际标题
for window in gw.getAllWindows():
    print(f"Window Title: {window.title}")

for window in gw.getAllWindows():
    if window.title == window_name:
        try:
            # 激活窗口
            window.activate()
            pyautogui.typewrite('111111')
            pyautogui.press('enter')
            print("输入成功")
            break
        except gw.PyGetWindowException:
            print("无法激活窗口")