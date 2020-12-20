import textwrap
ls = [1, "跑腿", "我這邊有一隻豬，徵求人幫我把它搬運到新體","100"]
print(ls[2])
content = textwrap.wrap(ls[2], 3)
print(type(content))
print(content)
text = "任務名稱：{}\n\n任務內容：".format(ls[1])
print(text)
for sentence in content:
    text += sentence
    text += "\n"
print(text)