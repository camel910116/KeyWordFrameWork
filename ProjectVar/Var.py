#encoding=utf-8
import os

#获取工程所在的目录的绝对路径
project_path=os.path.dirname(os.path.dirname(__file__))

#获取页面对象库文件的绝对路径
#page_object_repository_path=project_path.decode("utf-8")+u"/Conf/PageObjectRepository.ini"

#测试数据excel文件的绝对路径
test_data_excel_path = project_path.decode("utf-8")+u"/TestData/126邮箱的测试用例.xlsx"

#浏览器驱动所在的绝对路径
ieDriverFilePath="e:\\IEDriverServer"
chormeDriverFilePath="e:\\chromedriver"
firefoxDriverFilePath="e:\\geckodriver"

#设置截屏目录
screen_capturePicture_path = project_path+"\\ScreenPictures\\CapturePicture"
screen_errorPicture_path = project_path+"\\ScreenPictures\\ErrorPicture"

#excel列的含义
action_name_col_no =2
locator_method_col_no =3
locator_expression_col_no =4
action_value_col_no =5
action_elapse_time_col_no =7
action_result_col_no = 8
capture_screen_path_col_no =9
action_exception_col_no =10

if __name__=='__main__':
    print project_path
    #print page_object_repository_path
    print test_data_excel_path
    #print os.path.exists(page_object_repository_path)
    print os.path.exists(test_data_excel_path)
    print screen_capturePicture_path
    print screen_errorPicture_path
