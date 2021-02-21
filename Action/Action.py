#encoding=utf-8
from selenium import webdriver
import time
from Util.ObjectMap import *
from ProjectVar.Var import *
import traceback
from ProjectVar.Var import *
from Util.FormatTime import *
from Util.DirandFile import *

#定义全局浏览器driver变量
driver=None

def open_browser(browserName,*arg):
    #*arg传可变参数
    #**传字典，前面是key后面是value
    global driver
    try:
        if browserName.lower().strip()=="ie":
            driver = webdriver.Ie(executable_path = ieDriverFilePath)
        elif browserName.lower().strip()=="chrome":
            driver = webdriver.Ie(executable_path=chormeDriverFilePath)
        else:
            driver = webdriver.Ie(executable_path=firefoxDriverFilePath)
        #driver.get("https://www.baidu.com")
    except Exception,e:
        raise e

def visit_url(url,*arg):
    global driver
    try:
        driver.get(url)
    except Exception,e:
        raise e

def close_browser():
    global driver
    try:
        driver.quit()
    except Exception, e:
        raise e

def pause(seconds,*arg):
    time.sleep(float(seconds))

def switch_frame(locatorMethod,locatorExpression,*arg):
    global driver
    try:
        driver.switch_to.frame(getElement(driver, locatorMethod, locatorExpression))
    except Exception, e:
        raise e
        print "can not switch frame!"

def input_string(locatorMethod,locatorExpression,content,*arg):
    try:
        getElement(driver, locatorMethod, locatorExpression).clear()
        getElement(driver, locatorMethod, locatorExpression).send_keys(content)
    except Exception,e:
        raise e


def click(locatorMethod,locatorExpression,*arg):
    try:
        getElement(driver, locatorMethod, locatorExpression).click()
    except Exception,e:
        raise e

def login(usernameAndpassword,*arg):
    username,password =usernameAndpassword.split("||")
    open_browser("chrome")
    visit_url("http://mail.126.com")
    pause(3)
    switch_frame("id","x-URS-iframe")
    input_string("xpath", "//input[@name='email']", username)
    input_string("xpath", "//input[@name='password']", password)
    pause(3)
    click("id", "dologin")
    pause(3)

def add_contact_info(name,email,mobile,other_info):
    pause(3)
    click("xpath", "//div[text()='通讯录']")
    pause(3)
    click("xpath", "//span[text()='新建联系人']")
    input_string("xpath", "//a[@title='编辑详细姓名']/preceding-sibling::div/input", name)
    input_string("xpath", "//*[@id='iaddress_MAIL_wrap']//input", email)
    input_string("xpath", "//*[@id='iaddress_TEL_wrap']//dd//input", mobile)
    input_string("xpath", "//textarea", other_info)
    pause(3)
    click("xpath", "//span[text()='确 定']")


def assert_word(expected_word,*arg):
    try:
        assert True == (expected_word in driver.page_source)
    except AssertionError, e:
        raise e
    except Exception, e:
        raise e


def get_screen():
    global driver
    createDir(screen_capturePicture_path,dates())
    filePath = screen_capturePicture_path+"\\"+dates()+"\\"+times()+".jpg"
    print "filePath=",filePath
    try:
        driver.get_screenshot_as_file(filePath)
    except Exception,e:
        raise e
        print u"截屏失败"
    return filePath

def get_error_screen():
    global driver
    createDir(screen_errorPicture_path,dates())
    filePath = screen_errorPicture_path+"\\"+dates()+"\\"+times()+".jpg"
    print "filePath=",filePath
    try:
        driver.get_screenshot_as_file(filePath)
    except Exception,e:
        raise e
        print u"截屏失败"
    return filePath



if __name__=="__main__":

    #关键字框架和数据驱动框架相结合
    login_info = ("camel19910116||yan123", "camel19910213||yan123")
    contact_info = [("g1", "2ds30033@qq.com", "1393333333333", u"朋友"),
                    ("g2", "2ds30033@qq.com", "1393333333333", u"朋友")]
    for i in login_info:
        login(i)
        for j in contact_info:
            add_contact_info(j[0], j[1], j[2], j[3])
    # 等待10秒,以便登录成功后的页面加载完成
    # add_contact_info("gloryroad","2ds30033@qq.com","1393333333333",u"朋友")
    driver.quit()

    '''print screen_capturePicture_path
    login("camel19910116||yan123")
    assert_word(u"退出")
    add_contact_info(u"仇瑞平","xxx@qq.com","1391111111",u"朋友")
    get_screen(screen_capturePicture_path+"\\1.png")
    pause(3)'''












    '''click("xpath","//div[text()='通讯录']")
    pause(3)
    click("xpath", "//span[text()='新建联系人']")
    input_string("xpath","//a[@title='编辑详细姓名']/preceding-sibling::div/input",u"仇瑞平")
    input_string("xpath", "//*[@id='iaddress_MAIL_wrap']//input", "xxx@qq.com")
    input_string("xpath", "//*[@id='iaddress_TEL_wrap']//dd//input", "13911111")
    input_string("xpath", "//textarea", u"朋友")
    pause(3)
    click("xpath","//span[text()='确 定']")'''