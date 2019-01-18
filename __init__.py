import makePaper as mp
import commFunc as cf
import handleExcel as he

import logging,random,os
logging.basicConfig(level=logging.ERROR)

bankDest = [r'D:\文档\安规考试\附件4：秋查安规考试变电题库.xlsx',r'D:\文档\安规考试\附件5：秋查安规考试附加题库.xlsx']
#每套题库抽题数量，依次为单选、多选、判断、简答
bankSelect = [[29,15,39,0],[1,0,1,0]]

#样式文件地址
templeDest = r'D:\文档\安规考试\template.docx'
#各题分值,分别为单选、多选、判断、简答
quizMark = [1,2,1,0]
#题目序号
quizSeq = 1
#所有筛选好的试题
allQuizMatrix = []
#试题及答案保存地址
QuesPath = r'D:\文档\安规考试\交规'
#出试题份数
paperNum = 30

def countNum(bank,select):
    p = []
    q = []
    j = 0
    for i in bank:
        logging.info(i)
        sC,mC,jC,jQ = cf.calQueNum(cf.openQsBank(i))

        random.shuffle(sC)
        random.shuffle(mC)
        random.shuffle(jC)
        random.shuffle(jQ)

        select_new = select[j]
        j += 1

        p.append(sC[0:select_new[0]])
        p.append(mC[0:select_new[1]])
        p.append(jC[0:select_new[2]])
        p.append(jQ[0:select_new[3]])
        q.append(p)
        p = []

    return q

def wrtQuizChoice(bank,matrix,docQ,type):
    global quizSeq

    total = 0

    for i in range(len(bankSelect)):
        total += bankSelect[i][type]

    if type == 0:
        quesTitle = '一、单选题（共' + str(total) + '道题，每题' + str(quizMark[0]) + '分）'
    elif type == 1:
        quesTitle = '二、多选题（共' + str(total) + '道题，每题' + str(quizMark[1]) + '分）'

    mp.wrtQuesTitle(docQ,quesTitle)

    j = 0
    quizAns = []
    for i in bank:
        useSheet = cf.openQsBank(i)

        if matrix[j][type] != []:

            for n in matrix[j][type]:

                mp.wrtQues(docQ,quizSeq,cf.getCellValue(useSheet,n,2))
                quizSeq += 1
                quizChoice = []
                for num in range(3,7):
                    quizChoice.append(cf.getCellValue(useSheet,n,num))

                mp.wrtChoice(docQ,quizChoice)
                quizAns.append(cf.getCellValue(useSheet,n,7))

        j += 1

    return quizAns

def wrtOtherQuiz(bank,matrix,docQ,type):
    global quizSeq

    total = 0

    for i in range(len(bankSelect)):
        total += bankSelect[i][type]

    if type == 2:
        quesTitle = '三、判断题（共' + str(total) + '道题，每题' + str(quizMark[2]) + '分，正确涂A，错误涂B）'
    elif type == 3:
        if total != 0:
            quesTitle = '四、简答题（共' + str(total) + '道题，每题' + str(quizMark[3]) + '分）'
        elif total == 0:
            quesTitle = ''

    mp.wrtQuesTitle(docQ,quesTitle)

    j = 0
    quizAns = []
    for i in bank:
        useSheet = cf.openQsBank(i)

        if matrix[j][type] != []:

            for n in matrix[j][type]:

                mp.wrtQues(docQ,quizSeq,cf.getCellValue(useSheet,n,2))
                quizSeq += 1

                quizAns.append(cf.getCellValue(useSheet,n,7))

        j += 1

    return quizAns



def finalStep(num,templePath):

    for paperCount in range(1,num+1):

        # 创建试卷文件
        docQues = cf.creatDoc(templePath)
        # 创建答案文件
        docAns = cf.creatDoc(templePath)

        title = '2018年秋查规程考试交规第' + str(paperCount) + '套试题'

        mp.wrtTitle(docQues,title)
        mp.wrtTitle(docAns,title)

        allQuizMatrix = countNum(bankDest,bankSelect)

        ans_Single = wrtQuizChoice(bankDest,allQuizMatrix,docQues,0)
        ans_Multi = wrtQuizChoice(bankDest,allQuizMatrix,docQues,1)
        ans_jChoice = wrtOtherQuiz(bankDest,allQuizMatrix,docQues,2)
        ans_jQuiz = wrtOtherQuiz(bankDest,allQuizMatrix,docQues,3)

        k = mp.wrtAns(docAns,ans_Single,'单选题',0)
        k = mp.wrtAns(docAns,ans_Multi,'多选题',k)
        k = mp.wrtAns(docAns,ans_jChoice,'判断题',k)
        k = mp.wrtAns(docAns,ans_jQuiz,'简答题',k)

        Qfilename = '交规第' + str(paperCount) + '套试题.docx'
        Afilename = '交规第' + str(paperCount) + '套答案.docx'
        QuesDest = os.path.join(QuesPath,Qfilename)
        AnsDest = os.path.join(QuesPath,Afilename)

        docQues.save(QuesDest)
        docAns.save(AnsDest)

        print('Progressing['+ '*' * paperCount + ' '*(num - paperCount) + ']')

        global quizSeq
        quizSeq = 1

    return None

finalStep(paperNum,templeDest)
