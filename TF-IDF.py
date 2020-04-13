#!/usr/bin/env python
# coding: utf-8

# In[1]:


import openpyxl


# In[2]:


srcFilePath = r'C:\Users\hasee\Desktop\BJ_Text\AI＋精神用-自杀风险6-10级汇总（2018.11.06-2020.01.28）未剔除重复.xlsx'
srcFile = openpyxl.load_workbook(srcFilePath)
srcDatasheet = srcFile.worksheets[0]
srcDatasheet.delete_rows(0) # 删去表头


# In[3]:


#print(len(srcDatasheet["G"]))
#print(srcDatasheet["G2"].value)
srcTextList = srcDatasheet["G"]
srcTextList = [str(srcTextList[i].value) for i in range(len(srcTextList))]


# In[ ]:





# In[4]:


import jieba
import jieba.analyse


# In[5]:


#seg_list = jieba.cut(srcTextList[0])   #默认模式
#print([" ".join(seg_list) ])
tagTextList = [" ".join(jieba.cut(srcTextList[i])) for i in range(len(srcTextList))]
print(tagTextList[0])
print(tagTextList[1])


# In[39]:


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

from sklearn import feature_extraction  
from sklearn.feature_extraction.text import TfidfTransformer  
from sklearn.feature_extraction.text import CountVectorizer  
corpus=["我 来到 北京 清华大学",#第一类文本切词后的结果，词之间以空格隔开  
       "他 来到 了 网易 杭研 大厦",#第二类文本的切词结果  
        "小明 硕士 毕业 与 中国 科学院",#第三类文本的切词结果  
        "我 爱 北京 天安门"]#第四类文本的切词结果  
vectorizer=CountVectorizer()#该类会将文本中的词语转换为词频矩阵，矩阵元素a[i][j] 表示j词在i类文本下的词频  
transformer=TfidfTransformer()#该类会统计每个词语的tf-idf权值  
tfidf=transformer.fit_transform(vectorizer.fit_transform(corpus))#第一个fit_transform是计算tf-idf，第二个fit_transform是将文本转为词频矩阵  
word=vectorizer.get_feature_names()#获取词袋模型中的所有词语  
weight=tfidf.toarray()#将tf-idf矩阵抽取出来，元素a[i][j]表示j词在i类文本中的tf-idf权重  
for i in range(len(weight)):#打印每类文本的tf-idf词语权重，第一个for遍历所有文本，第二个for便利某一类文本下的词语权重  
    print (u"-------这里输出第",i,u"类文本的词语tf-idf权重------"  )
    for j in range(len(word)):  
        print (word[j],weight[i][j]  )
# In[26]:


print(len(tagDB[0]))
print(len(weightDB[0]))


# In[41]:


topN = 5

for sentence in range(20): # tarDB for full
    print(u"第", sentence, u"个语句的关键字：")
    for i in range(topN):
        print(tagDB[sentence][i], "\t:\t", weightDB[sentence][i])
    print("\n")


# In[ ]:




