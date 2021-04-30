from aip import AipImageClassify
from aip import AipImageSearch
from aip import AipOcr
import json
""" 你的 APPID AK SK """
APP_ID = '15790468'
API_KEY = 'CYVFFQZqyXkEQWo5dEzdBn21'
SECRET_KEY = 'PZFYzTpoOFcl9yG1thaxPW3X9zPGEfT8'


client = AipOcr(APP_ID, API_KEY, SECRET_KEY)

""" 读取图片 """
def get_file_content(filePath):
    with open(filePath, 'rb') as fp:
        return fp.read()

image = get_file_content(r'C:\Users\Administrator\Desktop\10574_scnuhuangli_1_100P1102404840769851392_建设进度表_6.jpg')

""" 调用通用文字识别（含位置高精度版） """
client.accurate(image);

""" 如果有可选参数 """
options = {}
options["recognize_granularity"] = "big"
options["detect_direction"] = "true"
options["vertexes_location"] = "true"
options["probability"] = "true"

""" 带参数调用通用文字识别（含位置高精度版） """
aa=client.accurate(image, options)
m=1;
print(aa.words_result);


