import re
mail = "b09704063@ntu.edu.tw"
data = re.findall(r"^[a-zA-Z0-9][\w\.-]*[a-zA-Z0-9]@ntu.edu.tw", mail)
print(data)