tplt = "{0:<8}      {1:{3}^10}    {2:>10}"
word = tplt.format("1", "跑腿", "20", chr(12288))
print(type(word))
print(word[0])