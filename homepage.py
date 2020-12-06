import tkinter as tk
import sys
import os
# from func1_png import img as func1
# from pageBackground_png import img as pageBackground
# import base64
"""
顏色代碼：
#283845 夜幕藍
#363636 微黑
#DF9A57 金色
文字大小為微軟正黑58粗體
"""



def login():
    """登入畫面"""
    # 背景頁建立
    global loginWin
    loginWin = tk.Frame(win)
    loginWin.config(width = 1024, height = 699, bg = "#363636")
    loginWin.place(x = 0, y = 0)

    # 標題建立
    global ntuCoinImg
    ntuCoinImg = tk.PhotoImage(file = "ntucoin.png")
    title = tk.Label(loginWin)
    title.config(image = ntuCoinImg, width = 450, height = 88)
    title.place(anchor = "n",x=512, y=50)

    # 用戶輸入欄
    userTitle = tk.Label(loginWin, text = "電郵地址或用戶名稱")
    userTitle.config(font = "微軟正黑體 14 bold", bg = "#363636", fg = "white")
    userTitle.place(anchor = "w", x = 300, y = 200)
    user = tk.Entry(loginWin)
    userText = tk.StringVar()
    user.config(textvariable = userText, width = 38, font = "arial 14")
    user.place(anchor = "w", x = 300, y = 230)

    # 密碼欄
    passwordTitle = tk.Label(loginWin, text = "密碼")
    passwordTitle.config(font = "微軟正黑體 14 bold", bg = "#363636", fg = "white")
    passwordTitle.place(anchor = "w", x = 300, y = 280)
    password = tk.Entry(loginWin)
    passwordText = tk.StringVar()
    password.config(textvariable = passwordText, width = 38, font = "arial 14", show = "●")
    password.place(anchor = "w", x = 300, y = 310)

    # 登入鍵
    global loginInImg
    loginInImg = tk.PhotoImage(file = "登入鍵.png")
    loginBtn = tk.Button(loginWin)
    loginBtn.config(image = loginInImg, relief = "flat", width = 450, height = 36)
    loginBtn.config(command = homepage, cursor = "hand2")
    loginBtn.place(anchor = "center", x = 512, y = 370)

    # 未註冊帳戶提示
    global notSignUpImg
    notSignUpImg = tk.PhotoImage(file = "未註冊帳戶.png")
    loginBtn = tk.Label(loginWin)
    loginBtn.config(image = notSignUpImg, relief = "flat", width = 450, height = 74)
    loginBtn.place(anchor = "center", x = 512, y = 483)

    # 註冊鍵
    global signUpImg
    signUpImg = tk.PhotoImage(file = "註冊鍵.png")
    loginBtn = tk.Button(loginWin)
    loginBtn.config(image = signUpImg, relief = "flat", width = 450, height = 36)
    loginBtn.config(command = signUp, cursor = "hand2")
    loginBtn.place(anchor = "center", x = 512, y = 550)

    # 離開鍵
    exitBtn = tk.Button(loginWin, text = "離開系統\nExit")
    exitBtn.config(font = "微軟正黑體 15 bold", bg = "#363636", fg = "white", relief = "flat")
    exitBtn.config(activebackground = "#363636", activeforeground = "#DF2935")
    exitBtn.config(command = whetherExit, cursor = "hand2")
    exitBtn.place(anchor = "se",x=1024, y=699)


def signUp():
    """註冊頁面"""
    # 背景頁建立
    signUpWin = tk.Frame(loginWin)
    signUpWin.config(width = 1024, height = 699, bg = "#363636")
    signUpWin.place(x = 0, y = 0)

    # 標題建立
    global ntuCoinImg
    ntuCoinImg = tk.PhotoImage(file = "ntucoin.png")
    title = tk.Label(signUpWin)
    title.config(image = ntuCoinImg, width = 450, height = 88)
    title.place(anchor = "n",x=512, y=20)

    sign = tk.Label(signUpWin, text = "以電子郵件地址註冊")
    sign.config(font = "微軟正黑體 20 bold", bg = "#363636", fg = "white")
    sign.place(anchor = "center", x = 512, y = 130)

    # 輸入電子信箱欄
    mailTitle = tk.Label(signUpWin, text = "您的電子郵件地址是什麼？")
    mailTitle.config(font = "微軟正黑體 14 bold", bg = "#363636", fg = "white")
    mailTitle.place(anchor = "w", x = 300, y = 170)
    mail = tk.Entry(signUpWin)
    mailText = tk.StringVar()
    mail.config(textvariable = mailText, width = 38, font = "arial 14")
    mail.place(anchor = "w", x = 300, y = 200)

    # 建立密碼
    passwordTitle = tk.Label(signUpWin, text = "建立密碼")
    passwordTitle.config(font = "微軟正黑體 14 bold", bg = "#363636", fg = "white")
    passwordTitle.place(anchor = "w", x = 300, y = 250)
    password = tk.Entry(signUpWin)
    passwordText = tk.StringVar()
    password.config(textvariable = passwordText, width = 38, font = "arial 14", show = "●")
    password.place(anchor = "w", x = 300, y = 280)

    # 確認密碼
    confirmTitle = tk.Label(signUpWin, text = "確認密碼")
    confirmTitle.config(font = "微軟正黑體 14 bold", bg = "#363636", fg = "white")
    confirmTitle.place(anchor = "w", x = 300, y = 330)
    confirm = tk.Entry(signUpWin)
    confirmText = tk.StringVar()
    confirm.config(textvariable = confirmText, width = 38, font = "arial 14", show = "●")
    confirm.place(anchor = "w", x = 300, y = 360)

    passwordWarning = tk.Label(signUpWin, text = "❕ 兩次密碼不相同")
    passwordWarning.config(font = "微軟正黑體 10", bg = "#363636", fg = "red")
    passwordWarning.place(anchor = "w", x = 300, y = 385)

    # 用戶名稱
    userTitle = tk.Label(signUpWin, text = "用戶名稱（12個字元以下）")
    userTitle.config(font = "微軟正黑體 14 bold", bg = "#363636", fg = "white")
    userTitle.place(anchor = "w", x = 300, y = 410)
    user = tk.Entry(signUpWin)
    userText = tk.StringVar()
    user.config(textvariable = userText, width = 38, font = "arial 14")
    user.place(anchor = "w", x = 300, y = 440)

    # username = user.get()
    # if len(username) > 12:
    nameWarning = tk.Label(signUpWin, text = "❕ 用戶名稱不符合規範")
    nameWarning.config(font = "微軟正黑體 10", bg = "#363636", fg = "red")
    nameWarning.place(anchor = "w", x = 300, y = 465)

    # 註冊鍵
    global signUpImg
    signUpImg = tk.PhotoImage(file = "註冊鍵.png")
    loginBtn = tk.Button(signUpWin)
    loginBtn.config(image = signUpImg, relief = "flat", width = 450, height = 36)
    loginBtn.config(command = login, cursor = "hand2")
    loginBtn.place(anchor = "center", x = 512, y = 520)

    username = userText.get()
    # if len(username) > 12:
    #     warning = tk.Label(signUpWin, text = "❕ 用戶名稱不符合規範")
    #     warning.config(font = "微軟正黑體 10", bg = "#363636", fg = "red")
    #     warning.place(anchor = "w", x = 300, y = 470)

    # 已經有帳戶了？
    haveAnAccountTitle = tk.Label(signUpWin, text = "已經有帳戶了？")
    haveAnAccountTitle.config(font = "微軟正黑體 22 bold", bg = "#363636", fg = "white")
    haveAnAccountTitle.place(anchor = "center", x = 512, y = 600)

    haveAnAccountBtn = tk.Button(signUpWin, text = "登入 > ")
    haveAnAccountBtn.config(font = "微軟正黑體 16 bold underline", bg = "#363636", fg = "#0066CC")
    haveAnAccountBtn.config(relief = "flat", activebackground = "#363636", activeforeground = "white")
    haveAnAccountBtn.config(command = login, cursor = "hand2")
    haveAnAccountBtn.place(anchor = "center", x = 512, y = 640)


def homepage():
    """主畫面"""
    global homeWin
    # 主畫面建立
    homeWin = tk.Frame(win)
    homeWin.config(width = 1024, height = 699, bg = "#363636")
    homeWin.place(x = 0, y = 0)


    # 標題建立
    global ntuCoinImg
    title = tk.PhotoImage(file = "ntucoin.png")
    title = tk.Label(homeWin)
    title.config(image = ntuCoinImg, width = 450, height = 88)
    title.place(anchor = "n",x=512, y=20)

    # 離開鍵建立
    exitBtn = tk.Button(homeWin, text = "離開系統\nExit")
    exitBtn.config(font = "微軟正黑體 15 bold", bg = "#363636", fg = "white", relief = "flat")
    exitBtn.config(activebackground = "#363636", activeforeground = "#DF2935")
    exitBtn.config(command = whetherExit, cursor = "hand2")
    # exitBtn.config(font = "微軟正黑體 15 bold", bg = "#DF9A57", fg = "black", relief = "solid", command = changeFullScreen)
    exitBtn.place(anchor = "se",x=1024, y=699)

    # 功能鍵建立
    global exchangeImg
    global taskImg
    global recordImg
    exchangeImg = tk.PhotoImage(file = "貨幣交換.png")
    taskImg = tk.PhotoImage(file = "任務.png")
    recordImg = tk.PhotoImage(file = "記錄查詢.png")

    exchange = tk.Button(homeWin)
    exchange.config(image = exchangeImg, width = 150, height = 400)
    exchange.config(relief = "flat")
    exchange.config(command = exchangeSys, cursor = "hand2")
    exchange.place(anchor = "w", x = 143, y = 350)

    task = tk.Button(homeWin)
    task.config(image = taskImg, width = 150, height = 400)
    task.config(relief = "flat")
    task.config(command = taskSys, cursor = "hand2")
    task.place(anchor = "w", x = 437, y = 350)

    record = tk.Button(homeWin)
    record.config(image = recordImg, width = 150, height = 400)
    record.config(relief = "flat")
    record.config(command = recordSys, cursor = "hand2")
    record.place(anchor = "w", x = 731, y = 350)

    # func4 = tk.Button(homeWin)
    # func4.config(image = func1Img, width = 150, height = 400, font = "微軟正黑體 30 bold")
    # func4.config(bg = "#DF9A57", fg = "black", relief = "flat")
    # func4.config(command = function4)
    # func4.place(anchor = "w", x = 786, y = 350)

    # 說明鍵建立
    helpBtn = tk.Button(homeWin, text = "功能說明\nHelp")
    helpBtn.config(font = "微軟正黑體 15 bold", bg = "#363636", fg = "white", relief = "flat")
    helpBtn.config(activebackground = "#363636", activeforeground = "#DF2935")
    helpBtn.config(command = helpPage, cursor = "hand2")
    helpBtn.place(anchor = "sw",x=0, y=699)


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

    # 確認是否離開畫面
    # whetherExitWin = tk.Label()
    # whetherExitWin.config(text = "您確定要離開系統嗎？\n", font = "微軟正黑體 40 bold", bg = "#99B2DD")
    # whetherExitWin.config(width = 18, height = 4)
    whetherExitWin = tk.Frame(win)
    whetherExitWin.config(width = 600, height = 300, bg = "#99B2DD")
    whetherExitWin.place(anchor = "center", x = 512, y = 350)

    whetherExitText = tk.Label(whetherExitWin)
    whetherExitText.config(text = "您確定要離開系統嗎？\n", font = "微軟正黑體 38 bold", bg = "#99B2DD")
    whetherExitText.config(width = 16, height = 8)
    whetherExitText.place(anchor = "center", x= 300, y = 120)

    # YES鍵
    yesBtn = tk.Button(whetherExitWin, text = "YES")
    yesBtn.config(font = "微軟正黑體 30 bold", bg = "#F3BB91", fg = "red", activebackground = "yellow")
    yesBtn.config(relief = "flat")
    yesBtn.config(command = exit, cursor = "hand2")
    yesBtn.place(anchor = "center", x = 200, y = 200)
    
    # NO鍵
    noBtn = tk.Button(whetherExitWin, text = "NO")
    noBtn.config(font = "微軟正黑體 30 bold", bg = "#F3BB91", fg = "blue", activebackground = "yellow")
    noBtn.config(relief = "flat")
    noBtn.config(command = whetherExitWin.destroy, cursor = "hand2")
    noBtn.place(anchor = "center", x = 390, y = 200)

    whetherExitWin.place(anchor = "center", x = 512, y = 350)


def helpPage():
    """說明頁"""
    # 背景頁建立
    helpWin = tk.Frame(homeWin)
    helpWin.config(width = 1024, height = 699, bg = "#363636")
    helpWin.place(x = 0, y = 0)

    # 返回鍵建立
    backBtn = tk.Button(helpWin, text = "回到首頁\nBack")
    backBtn.config(font = "微軟正黑體 15 bold", bg = "#363636", fg = "white", relief = "flat")
    backBtn.config(activebackground = "#363636", activeforeground = "#DF2935")
    backBtn.config(command = helpWin.destroy, cursor = "hand2")
    backBtn.place(anchor = "se",x=1024, y=699)

    # 說明文字建立
    helpText1 = tk.Label(helpWin, text = "  NTU Coin")
    helpText1.config(font = "微軟正黑體 45 bold", bg = "#363636", fg = "white")

    helpText2 = tk.Label(helpWin, text = """
    NTU Coin 使得臺大教職員生交易更便利！舉凡是\n
    臺大內部需要付錢的地方（學餐、小福、便利商店\n
    、停車費、影印費、販賣機等等），皆可使用NTU \n
    Coin 進行支付。同時，大多數臺大周圍的餐廳、\n
    店家也接受NTU Coin。NTU Coin 由臺大校方負責\n
    發行及兌換，以確保NTU Coin的可流通性。
    """)
    helpText2.config(font = "微軟正黑體 24 bold", bg = "#363636", fg = "white", justify = "left")

    helpText1.place(x = 0, y = 20)
    helpText2.place(x = 0, y = 100)


def exchangeSys():
    """貨幣交換系統"""
    # 背景頁建立
    exchangeSysWin = tk.Frame(homeWin)
    exchangeSysWin.config(width = 1024, height = 699, bg = "#363636")
    exchangeSysWin.place(x = 0, y = 0)

    # 返回鍵建立
    backBtn = tk.Button(exchangeSysWin, text = "回到首頁\nBack")
    backBtn.config(font = "微軟正黑體 15 bold", bg = "#363636", fg = "white", relief = "flat")
    backBtn.config(activebackground = "#363636", activeforeground = "#DF2935")
    backBtn.config(command = exchangeSysWin.destroy, cursor = "hand2")
    backBtn.place(anchor = "se",x=1024, y=699)


def taskSys():
    """任務系統"""
    # 說明頁建立
    taskSysWin = tk.Frame(homeWin)
    taskSysWin.config(width = 1024, height = 699, bg = "#363636")
    taskSysWin.place(x = 0, y = 0)

    # 返回鍵建立
    backBtn = tk.Button(taskSysWin, text = "回到首頁\nBack")
    backBtn.config(font = "微軟正黑體 15 bold", bg = "#363636", fg = "white", relief = "flat")
    backBtn.config(activebackground = "#363636", activeforeground = "#DF2935")
    backBtn.config(command = taskSysWin.destroy, cursor = "hand2")
    backBtn.place(anchor = "se",x=1024, y=699)


def recordSys():
    """記錄系統"""
    # 說明頁建立
    recordSysWin = tk.Frame(homeWin)
    recordSysWin.config(width = 1024, height = 699, bg = "#363636")
    recordSysWin.place(x = 0, y = 0)

    # 返回鍵建立
    backBtn = tk.Button(recordSysWin, text = "回到首頁\nBack")
    backBtn.config(font = "微軟正黑體 15 bold", bg = "#363636", fg = "white", relief = "flat")
    backBtn.config(activebackground = "#363636", activeforeground = "#DF2935")
    backBtn.config(command = recordSysWin.destroy, cursor = "hand2")
    backBtn.place(anchor = "se",x=1024, y=699)


def function4():
    """功能四"""
    # 說明頁建立
    func4Win = tk.Frame(homeWin)
    func4Win.config(width = 1024, height = 699, bg = "#363636")
    func4Win.place(x = 0, y = 0)

    # 返回鍵建立
    backBtn = tk.Button(func4Win, text = "回到首頁\nBack")
    backBtn.config(font = "微軟正黑體 15 bold", bg = "#363636", fg = "white", relief = "flat")
    backBtn.config(activebackground = "#363636", activeforeground = "#DF2935")
    backBtn.config(command = func4Win.destroy)
    backBtn.place(anchor = "se",x=1024, y=699)


# 方便pyinstaller將圖片攜帶
# picData = open('func1.png', 'wb')
# picData.write(base64.b64decode(func1))
# picData.close()
# picData = open('pageBackground.png', 'wb')
# picData.write(base64.b64decode(pageBackground))
# picData.close()

# 建立視窗
win = tk.Tk()
win.title("NTU Coin")
state = False
win.geometry("1024x699+172+0")
win.resizable(False, False)
win.config(bg = "#363636")
login()
win.mainloop()
