def checkSignUpInfo():
    """檢查輸入的資訊是否符合規則"""
    mail = mailText.get()
    password1 = passwordText.get()
    password2 = confirmText.get()
    userName = userText.get()
    state = []  # 最後用來判斷有無錯誤發生

    # 檢查信箱欄
    mW = tk.StringVar()
    mW.set("")
    mailWarning = tk.Label(signUpWin, textvariable = mW)
    mailWarning.config(font = "微軟正黑體 10", bg = "#363636", fg = "red")
    mailWarning.place(anchor = "w", x = 300, y = 225)

    # 檢查密碼欄
    p1W = tk.StringVar()
    p1W.set("")
    password1Warning = tk.Label(signUpWin, textvariable = p1W)
    password1Warning.config(font = "微軟正黑體 10", bg = "#363636", fg = "red")
    password1Warning.place(anchor = "w", x = 300, y = 305)

    # 檢查密碼確認欄
    p2W = tk.StringVar()
    p2W.set("")
    password2Warning = tk.Label(signUpWin, textvariable = p2W)
    password2Warning.config(font = "微軟正黑體 10", bg = "#363636", fg = "red")
    password2Warning.place(anchor = "w", x = 300, y = 385)
    
    # 檢查用戶名稱欄
    nW = tk.StringVar()
    nW.set("")
    nameWarning = tk.Label(signUpWin, textvariable = nW)
    nameWarning.config(font = "微軟正黑體 10", bg = "#363636", fg = "red")
    nameWarning.place(anchor = "w", x = 300, y = 465)

    # 印出錯誤信息
    if mail == "":
        mW.set("❕ 此欄不得為空                     ")
    else:
        data = re.findall(r"^[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]@[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]\.[a-zA-Z][a-zA-Z\.]*[a-zA-Z]$", mail)
        if len(data) != 1:
            mW.set("❕ 電郵地址不符合格式                       ")
        else:
            mW.set("                                            ")
            state.append(0)

    if password1 == "":
        p1W.set("❕ 此欄不得為空                        ")
    else:
        p1W.set("                                           ")
        state.append(0)

    if password2 == "":
        p2W.set("❕ 此欄不得為空                        ")
    else:
        if password1 != password2:
            p2W.set("❕ 兩次密碼不相同                      ")
        else:
            p2W.set("                                           ")
            state.append(0)

    if userName == "":
        nW.set("❕ 此欄不得為空                     ")
    else: 
        if len(userName) > 12:
            nW.set("❕ 用戶名稱不符合規範                       ")
        else:
            nW.set("                                            ")
            state.append(0)
    if len(state) == 4: # 若回傳四個0即代表四個欄位都符合規格
        login()


