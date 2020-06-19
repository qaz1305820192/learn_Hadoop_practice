

# -*- coding: utf-8 -*-
"""
Created on Thu Jun 18 16:39:55 2020
查找多个doc中的某些关键字，将含有这些关键字的word文件进行列表显示
@author: 13058
"""

from win32com import client as wc
import docx
import os


   
def doSaveAas(doc_path):
    docx_path = doc_path.replace("doc", "docx")
    word = wc.Dispatch('Word.Application')
    doc = word.Documents.Open(doc_path)# 目标路径下的文件
    doc.SaveAs(docx_path, 12, False, "", True, "", False, False, False, False)# 转化后路径下的文件
    doc.Close()
    return docx_path

#获取文件列表
def pathsDoc(path):
        path_collection=[]
        for dirpath,dirnames,filenames in os.walk(path):
                for file in filenames:
                        fullpath=os.path.join(dirpath,file)
                        if fullpath.endswith(".doc"):
                            path_collection.append(fullpath)
        return path_collection
#获取文件列表
def pathsDocx(path):
        path_collection=[]
        for dirpath,dirnames,filenames in os.walk(path):
                for file in filenames:
                        fullpath=os.path.join(dirpath,file)
                        if fullpath.endswith(".docx"):
                            path_collection.append(fullpath)
        return path_collection


def readDocx(filepath):
    #获取文档对象
    file=docx.Document(filepath)
    content=""
    for para in file.paragraphs:
        content=content+para.text
    return content
 

def mathcKey(filepath,strKey):
    content=readDocx(filepath)
    if content.find(strKey, 0, len(content))!=-1:
        return path
    return ""


def resultFile(fileDir,keyWord):
    path=fileDir
    #首先把doc文件转化成docx文件
    for filepath in pathsDoc(path):
        doSaveAas(filepath)
    #保存查到的文件
    result=[]
    for filepath in pathsDocx(path):
        #如果匹配到了则保存到result中
        if mathcKey(filepath,keyWord)!='':
            result.append(filepath)
    #删除程序生成的docx文件
    for filepath in pathsDoc(path):
        os.remove(filepath+'x')
    #输出结果
    result2=[]        
    for str1 in result:
        print(str1)
        #去掉文件结尾的 x
        result2.append(str1[0:len(str1)-1])
    return result2        
    
from shutil import copyfile     
import easygui        
if __name__ == '__main__':
    #指定查找的目录
    
   # doSaveAas("C:\\Users\\13058\\Desktop\\pp\\70邱共和张勇等走私普通货物物品罪二审刑事裁定书.doc")
    
    if easygui.ynbox('选择要搜索的文件夹', '亲爱的小彭', ('是', '否')):
        fileDir=easygui.diropenbox()
        keyWord=easygui.enterbox("要搜索的关键字",title="要搜索的关键字", default="", strip=True, image=None, root=None)
        
        resultfile=resultFile(fileDir,keyWord)
        if len(resultfile)!=0:
            if False == os.path.exists(path+"\查找的结果"):
                os.mkdir(fileDir+'\查找的结果')
            for file in resultfile:
                filename=file.split('\\')[len(file.split('\\'))-1]
                copyfile(file,fileDir+'\查找的结果\\'+filename)
        else:
            easygui.msgbox("这个文件里的doc文件没有包含："+keyWord)
                
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
        
