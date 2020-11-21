import tkinter as tk
import sys
"""
顏色代碼：
#283845 夜幕藍
#363636 微黑
"""
def changeFullScreen():
    """切換全螢幕"""
    global state
    state = not state
    if state == False:
        win.attributes("-fullscreen", state)
        win.geometry("1024x768+172+0")
        exitBtn.config(text = "打開全螢幕\nExpand")
    else:
        win.attributes("-fullscreen", state)
        exitBtn.config(text = "退出全螢幕\nExit")


def exit():
    """離開系統"""
    sys.exit()


# 主畫面建立
win = tk.Tk()
win.title("記帳小幫手")
state = False
win.geometry("1024x699+172+0")
win.resizable(False, False)
win.config(bg = "#DF9A57")

# 離開鍵建立
exitBtn = tk.Button(win, text = "離開系統\nExit")
exitBtn.config(font = "微軟正黑體 15 bold", bg = "#363636", fg = "white", relief = "flat", activebackground = "#DF2935")
exitBtn.config(command = exit)
# exitBtn.config(font = "微軟正黑體 15 bold", bg = "#DF9A57", fg = "black", relief = "solid", command = changeFullScreen)
exitBtn.place(anchor = "se",x=1024, y=699)

func1Img = tk.PhotoImage(file="func1.png") 
# func1Img = "C:\\Users\\User\\Desktop\\專案圖.jpg"
func1 = tk.Button()
func1.config(image = func1Img, width = 150, height = 400, font = "微軟正黑體 30 bold")
func1.config(bg = "#DF9A57", fg = "black", relief = "flat")
func1.place(anchor = "w", x = 84, y = 350)

func1 = tk.Button()
func1.config(image = func1Img, width = 150, height = 400, font = "微軟正黑體 30 bold")
func1.config(bg = "#DF9A57", fg = "black", relief = "flat")
func1.place(anchor = "w", x = 318, y = 350)

func1 = tk.Button()
func1.config(image = func1Img, width = 150, height = 400, font = "微軟正黑體 30 bold")
func1.config(bg = "#DF9A57", fg = "black", relief = "flat")
func1.place(anchor = "w", x = 552, y = 350)

func1 = tk.Button()
func1.config(image = func1Img, width = 150, height = 400, font = "微軟正黑體 30 bold")
func1.config(bg = "#DF9A57", fg = "black", relief = "flat")
func1.place(anchor = "w", x = 786, y = 350)

win.mainloop()
