#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import openpyxl


# In[ ]:


srcFilePath = r'C:\Users\hasee\Desktop\BJ_Text\AI＋精神用-自杀风险6-10级汇总（2018.11.06-2020.01.28）未剔除重复.xlsx'
srcFile = openpyxl.load_workbook(srcFilePath)
srcDatasheet = srcFile.worksheets[0]
srcDatasheet.delete_rows(0) # 删去表头


# In[ ]:


#print(len(srcDatasheet["G"]))
#print(srcDatasheet["G2"].value)
srcTextList = srcDatasheet["G"]
srcTextList = [str(srcTextList[i].value) for i in range(len(srcTextList))]


# In[ ]:


import jieba
import jieba.analyse


# In[ ]:


#seg_list = jieba.cut(srcTextList[0])   #默认模式
#print([" ".join(seg_list) ])
tagTextList = [" ".join(jieba.cut(srcTextList[i])) for i in range(len(srcTextList))]
print(tagTextList[0])
print(tagTextList[1])


# In[ ]:


from sklearn import feature_extraction
from sklearn.feature_extraction.text import CountVectorizer
from sklearn.feature_extraction.text import TfidfTransformer

stopWords = ["nbsp"]
vectorizer = CountVectorizer(stop_words=stopWords)
transformer=TfidfTransformer()
tf = vectorizer.fit_transform(tagTextList)
tfidf = transformer.fit_transform(tf)
tag = vectorizer.get_feature_names()
weight = tfidf.toarray()

resultMap = []
tagDB = []
weightDB = []
for i in range(len(weight)):
    sortedWeight, sortedTag = zip(*sorted(zip(weight[i],tag), reverse=True))
    tagDB.append(sortedTag)
    weightDB.append(sortedWeight)


# In[ ]:


print(len(tagDB[0]))
print(len(weightDB[0]))


# In[ ]:


topN = 5

for sentence in range(20): # tarDB for full
    print(u"第", sentence, u"个语句的关键字：")
    for i in range(topN):
        print(tagDB[sentence][i], "\t:\t", weightDB[sentence][i])
    print("\n")


# In[ ]:




