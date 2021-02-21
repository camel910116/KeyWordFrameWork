#encoding=utf-8
from ConfigParser import ConfigParser
from ProjectVar.Var import page_object_repository_path
#这个类主要是用来解析配置文件，从配置文件中读取出对象的定位方式和表达式

class ParsePageObjectRepositoryConfig(object):
    def __init__(self):
        self.cf=ConfigParser()
        self.cf.read(page_object_repository_path)

    def getItemsFromSection(self,sectionName):
        print self.cf.items(sectionName)
        return dict(self.cf.items(sectionName))

    def getOptionValue(self,sectionName,optionName):
        return self.cf.get(sectionName,optionName)

if __name__=="__main__":
    pp=ParsePageObjectRepositoryConfig()
    print pp.getItemsFromSection("126mail_login")
    print pp.getOptionValue("126mail_login","loginPage.username")
