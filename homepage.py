import tkinter as tk
import sys
import os
"""
顏色代碼：
#283845 夜幕藍
#363636 微黑
#DF9A57 金色
"""
def homepage():
    """主畫面"""
    # 主畫面建立
    win = tk.Tk()
    win.title("記帳小幫手")
    state = False
    win.geometry("1024x699+172+0")
    win.resizable(False, False)
    win.config(bg = "#363636")

    # 標題建立
    title = tk.Label(win, text = "         記 帳 小 幫 手         ")
    title.config(font = "微軟正黑體 40 bold", bg = "#363636", fg = "white")
    title.pack()

    # 離開鍵建立
    exitBtn = tk.Button(win, text = "離開系統\nExit")
    exitBtn.config(font = "微軟正黑體 15 bold", bg = "#363636", fg = "white", relief = "flat")
    exitBtn.config(activebackground = "#363636", activeforeground = "#DF2935")
    exitBtn.config(command = whetherExit)
    # exitBtn.config(font = "微軟正黑體 15 bold", bg = "#DF9A57", fg = "black", relief = "solid", command = changeFullScreen)
    exitBtn.place(anchor = "se",x=1024, y=699)

    # 功能鍵建立
    func1Img = tk.PhotoImage(file="func1.png") 
    func1 = tk.Button(win)
    func1.config(image = func1Img, width = 150, height = 400, font = "微軟正黑體 30 bold")
    func1.config(bg = "#DF9A57", fg = "black", relief = "flat")
    func1.place(anchor = "w", x = 84, y = 350)

    func1 = tk.Button(win)
    func1.config(image = func1Img, width = 150, height = 400, font = "微軟正黑體 30 bold")
    func1.config(bg = "#DF9A57", fg = "black", relief = "flat")
    func1.place(anchor = "w", x = 318, y = 350)

    func1 = tk.Button(win)
    func1.config(image = func1Img, width = 150, height = 400, font = "微軟正黑體 30 bold")
    func1.config(bg = "#DF9A57", fg = "black", relief = "flat")
    func1.place(anchor = "w", x = 552, y = 350)

    func1 = tk.Button(win)
    func1.config(image = func1Img, width = 150, height = 400, font = "微軟正黑體 30 bold")
    func1.config(bg = "#DF9A57", fg = "black", relief = "flat")
    func1.place(anchor = "w", x = 786, y = 350)

    # 說明鍵建立
    helpBtn = tk.Button(win, text = "功能說明\nHelp")
    helpBtn.config(font = "微軟正黑體 15 bold", bg = "#363636", fg = "white", relief = "flat")
    helpBtn.config(activebackground = "#363636", activeforeground = "#DF2935")
    helpBtn.place(anchor = "sw",x=0, y=699)

    win.mainloop()


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


def whetherExit():
    """詢問是否離開"""
    whetherExitWin = tk.Label()
    whetherExitWin.config(text = "您確定要離開系統嗎？\n", font = "微軟正黑體 40 bold", bg = "#99B2DD")
    whetherExitWin.config(width = 18, height = 4)

    yesBtn = tk.Button(whetherExitWin, text = "YES")
    yesBtn.config(font = "微軟正黑體 30 bold", bg = "#F3BB91", fg = "red", activebackground = "yellow")
    yesBtn.config(relief = "flat")
    yesBtn.config(command = exit)
    yesBtn.place(anchor = "center", x = 200, y = 200)

    noBtn = tk.Button(whetherExitWin, text = "NO")
    noBtn.config(font = "微軟正黑體 30 bold", bg = "#F3BB91", fg = "blue", activebackground = "yellow")
    noBtn.config(relief = "flat")
    noBtn.config(command = whetherExitWin.destroy)
    noBtn.place(anchor = "center", x = 390, y = 200)

    whetherExitWin.place(anchor = "center", x = 512, y = 350)


homepage()