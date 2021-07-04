from selenium import webdriver
import time
from bs4 import BeautifulSoup
import pyperclip #python实现复制粘贴
from pykeyboard import PyKeyboard
def main():
    browser = webdriver.Chrome()
    browser.get('https://baidu.com')  # 在当前浏览器中访问百度
    pyperclip.copy("selenium")
    browser.find_element_by_name('wd').click() #点击一下百度的输入框
    time.sleep(0.5)
    k = PyKeyboard()
    #模拟键盘点击ctrl+v
    k.press_key(k.control_key)
    k.tap_key('v')
    k.release_key(k.control_key)
    browser.find_element_by_name('wd').click()
    time.sleep(1000)
if __name__ == '__main__':
    main()
