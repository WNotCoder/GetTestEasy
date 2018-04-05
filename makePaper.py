
import os,docx,openpyxl,random,re
import logging
logging.basicConfig(level=logging.INFO)

import commFunc as cf

#写标题
def wrtTitle(doc,paperTitle):
    para = doc.add_paragraph(paperTitle)
    para.style = 'paperStyleTitle'
    return None

#写题型标题
def wrtQuesTitle(doc,quesTitle):
    para = doc.add_paragraph(quesTitle)
    para.style = 'paperStyleTitle'
    return None

#写题目
def wrtQues(doc,quesNO,quesContent):
    paraObj = doc.add_paragraph(str(quesNO) + chr(46) + quesContent)
    paraObj.style = 'paperStyleQues'
    return paraObj

#写选项
def wrtChoice(doc, quesCh):
    noCChoice = 0
    noDChoice = 0
    paraObj1 = doc.add_paragraph(quesCh[0])
    paraObj1.style = 'paperStyleOption'
    if len(quesCh[1]) < 15:
        paraObj1.add_run('    '+ quesCh[1])
    else:
        paraObj3 = doc.add_paragraph(quesCh[1])
        paraObj3.style = 'paperStyleOption'
    try:
        len(quesCh[2])
        paraObj2 = doc.add_paragraph(quesCh[2])
        paraObj2.style = 'paperStyleOption'
    except TypeError:
        noCChoice = 1
        noDChoice = 1
        return noCChoice,noDChoice

    try:
        if len(quesCh[3]) < 15:
            paraObj2.add_run('    '+ quesCh[3])
        else:
            paraobj4 = doc.add_paragraph(quesCh[3])
            paraobj4.style = 'paperStyleOption'
    except TypeError or NameError:
        noDChoice = 1
    return noCChoice,noDChoice


#写答案
def wrtAns(doc, quesAns, quesType,k):
    paraObj1 = doc.add_paragraph(quesType)
    paraObj1.style = 'paperStyleItem'
    paraObj2 = doc.add_paragraph('')
    paraObj2.style = 'paperStyleOption'
    for i in quesAns:
        k += 1
        paraObj2.add_run(str(k) + chr(46) + i + chr(32))
    return k



