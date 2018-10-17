import excelToolUtils
import wordCount
import time
import os
 
# 文件夹路径
filePath = "D:/测试文件/给彪哥.xlsx"
# 取前词数
wordIndex = 2


# 读取数据
messageDict = excelToolUtils.readExcel(filePath)
dateTime = time.strftime("%Y%m%d-%H%M%S", time.localtime())
# 接吧分词
for key in messageDict.keys():
    currenctWordIndex = 0
    words = wordCount.wordFrequencyCount(messageDict[key])

    # 统计词频
    dict = {}
    for word in words :
        if len(word) > 1 :
            if word in dict:
                dict[word] = dict[word] + 1
            else:
                dict[word] = 1
    # 对字典排序
    dict2 = sorted(dict,reverse=True)
    print("统计的分词后的词频")
    print(str(dict))
    print("根据词频排序后的结果")
    print(dict2)

    # 生成每行的数据
    rows = [["词组","词频"]]
    for i in dict2 :
        if currenctWordIndex == wordIndex:
            break
        row = []
        row.append(i)
        row.append(dict[i])
        rows.append(row)
        currenctWordIndex += 1
    print(rows)

    filePath="D:/数据处理结果/"+dateTime
    # 判断文件夹是否存在
    if os.path.exists(filePath):
        print("保存文件到="+filePath)
    else:
        os.makedirs(filePath)
        print("保存文件到=" + filePath)
    excelToolUtils.writeExcel(filePath+"/"+key+".xlsx","Sheet1",rows)

