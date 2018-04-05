import logging
logging.basicConfig(level=logging.INFO)

import openpyxl,re
import commFunc as cf

#删除空格
def cleanSpace(s):
    spacePattern = re.compile(r'\s')
    return spacePattern.sub('', s)

#删除序号项
def cleanNum(s):
    numPattern = re.compile(r'^[0-9]{1,3}')
    s = numPattern.sub('', s)
    commaPattern = re.compile(r'^([.、])')
    s = commaPattern.sub('', s)
    return  s

#删除答案项
def cleanAns(s):
    ansPattern = re.compile(r'[(（][ABCD ]{1,4}[)）]')
    return ansPattern.sub(r'(  )' , s)

#获取试题答案
def getAns(s):
    pattern = re.compile(r'(\S)*([(（])([ABCD ]{1,4})([)）])(\S)*')
    s = pattern.search(s)
    titleAns = s.group(3)
    return titleAns

#异常检测
def checkNoneType(content,rowNo,columnNo):
    rowNo = int(rowNo)
    columnNo = int(columnNo)
    if content == None or content.isspace():
        return True
    else:
        return False

def checkChoiceMark(content):
    pattern = r'A|B|C|D'
    if re.search(pattern,content) == None:
        return True
    else:
        return False

#形成标准题库
def standTitle(useSheet,titleNo):
    for i in titleNo:
        #删除空格
        for j in range(2,7):
            s = cleanSpace(cf.getCellValue(useSheet,i,j))
            cf.setCellValue(useSheet,i,j,s)
        #删除序号
            cf.setCellValue(useSheet,i,j,cleanNum(cf.getCellValue(useSheet,i,j)))

        #获取答案
        cf.setCellValue(useSheet,i,7,getAns(cf.getCellValue(useSheet,i,1)))
        #删除答案
        cf.setCellValue(useSheet,i,1,cleanAns(cf.getCellValue(useSheet,i,1)))
        return None




