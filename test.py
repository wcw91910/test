tplt_header = "{0:{4}<4}  {1:{4}^6}  {2:{4}^8}{3:{4}>6}"
tplt = "{0:<8}  {1:^12}  {2:{4}^8}{3:>12}"
header = ["日期", "用戶名稱", "內容", "金額"]
ls = [["12/11", "daniel", "吃飯", "$100"], ["5/12", "llllllllll", "幫忙搶課課課課課", "$1000"], ["12/13", "dog", "幫忙簽到", "$10"]]
print(tplt_header.format(header[0], header[1], header[2], header[3], chr(12288)))
for i, content in enumerate(ls):
    print(tplt.format(content[0], content[1], content[2], content[3], chr(12288)))