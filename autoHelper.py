from selenium import webdriver
from selenium.webdriver.common.by import By
import time

#自动化通用程序
class autoHelper(object):
    def __init__(self,url):    
        self.wd = webdriver.Chrome()
        self.url = url

    #通过xpath查询元素是否存在
    def is_element_exist(self,xpath):
        self.wd.implicitly_wait(0)
        ele = self.wd.find_elements(By.XPATH,xpath)
        self.wd.implicitly_wait(5)
        if	len(ele)==0:
            return False
        if len(ele)==1:
            return True
        else:
            return False 

    #点击确定
    def confirm_click(self):
        elements = self.wd.find_elements(By.CLASS_NAME,'u-form-btn-txt')
        for element in elements:
            if(element.text == '确定'):
                element.click()
                break

    #是否加载完成
    def load_complete(self,xpath):
        while True:
            ele=self.wd.find_elements(By.XPATH,xpath)
            if len(ele)==1:
                #print(ele[0].get_attribute("style")[0:13])
                if(ele[0].get_attribute("style")[0:13] =='display: none'):
                    break

    #切换到其他的界面
    def widowChang(self,name):
        for handle in self.wd.window_handles:
            # 先切换到该窗口
            self.wd.switch_to.window(handle)
            # 得到该窗口的标题栏字符串，判断是不是我们要操作的那个窗口
            if name in self.wd.title:
                # 如果是，那么这时候WebDriver对象就是对应的该该窗口，正好，跳出循环，
                break

    #外汇管理局登录界面MMTT
    def load(self):
        #url = 'http://zwfw.safe.gov.cn/asone/servlet/UniLoginServlet?COLLCC=1430122602&'
        self.wd.get(self.url)
        print(self.wd.switch_to.alert.text)
        # 点击 OK 按钮 
        self.wd.switch_to.alert.accept()
        mainWindow = self.wd.current_window_handle
        self.widowChang('弹出')
        self.wd.close()
        self.wd.switch_to.window(mainWindow)
        element_1 = self.wd.find_element(By.XPATH,'//*[@id="inputOrgCode"]')
        #公司代码
        element_1.send_keys('')
        element_1 = self.wd.find_element(By.XPATH,'//*[@id="inputUserCode"]')
        #账号
        element_1.send_keys('')
        element_1 = self.wd.find_element(By.XPATH,'//*[@id="pwdw"]')
        #密码
        element_1.send_keys('')
        element_1 = self.wd.find_element(By.XPATH,'//*[@id="tanchuang"]/table/tbody/tr/td[2]/img')
        element_1.click()
        login_button = self.wd.find_element(By.XPATH,'//*[@id="loginc"]')
        element_2 = self.wd.find_element(By.XPATH,'//*[@id="check"]')
        print(element_2.get_attribute('value'))
        while(True):
            if len(element_2.get_attribute('value')) == 4:
                login_button.click()
                break
            time.sleep(0.01)

    #外汇新增登录贸易
    def load_in(self):
        #url = 'http://zwfw.safe.gov.cn/asone/servlet/UniLoginServlet?COLLCC=1430122602&'
        self.wd.get(self.url)
        print(self.wd.switch_to.alert.text)
        # 点击 OK 按钮 
        self.wd.switch_to.alert.accept()
        mainWindow = self.wd.current_window_handle
        self.widowChang('弹出')
        self.wd.close()
        self.wd.switch_to.window(mainWindow)
        element_1 = self.wd.find_element(By.XPATH,'//*[@id="inputOrgCode"]')
        #公司代码
        element_1.send_keys('')
        element_1 = self.wd.find_element(By.XPATH,'//*[@id="inputUserCode"]')
        #账号
        element_1.send_keys('')
        element_1 = self.wd.find_element(By.XPATH,'//*[@id="pwdw"]')
        #密码
        element_1.send_keys('')
        element_1 = self.wd.find_element(By.XPATH,'//*[@id="tanchuang"]/table/tbody/tr/td[2]/img')
        element_1.click()
        login_button = self.wd.find_element(By.XPATH,'//*[@id="loginc"]')
        element_2 = self.wd.find_element(By.XPATH,'//*[@id="check"]')
        print(element_2.get_attribute('value'))
        while(True):
            if len(element_2.get_attribute('value')) == 4:
                login_button.click()
                break
            time.sleep(0.01)
