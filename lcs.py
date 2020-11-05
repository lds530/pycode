
import win32com
from win32com.client import Dispatch, constants
import os,sys

#LCS比较文档不同
class DiffLcs:
    
    #找两字串最长公共序列
    #str1 源串 #str2比较串，s:开始位置,e1，源串结束位置，e2比较串结束位置
    @staticmethod
    def genPartialDiff(str1,str2,s,e1,e2):
        len1=e1-s+2
        len2=e2-s+2
        
        #二维表初始化 row 为源串，col为比较串 [方向，值]
        tb=[[['',0] for j in  range(len2)] for i in  range(len1)]
        for i in range(1,len1):
            tb[i][0][0]=str1[i+s-1]
        for j in range(1,len2):
            tb[0][j][0]=str2[j+s-1]

        for i in range(1,len1):
            for j in range(1,len2):
                #左上↖
                if str1[i+s-1]==str2[j+s-1]:
                    tb[i][j]=['u',tb[i-1][j-1][1]+1]
                #左←
                elif tb[i][j-1][1]>tb[i-1][j][1]:
                    tb[i][j]=['i',tb[i][j-1][1]]
                #上↑
                else:
                    tb[i][j]=['d',tb[i-1][j][1]]
        '''
        for a in tb:
            print(a)
        '''

        i=len1-1
        j=len2-1
        #回溯[{u|d|i},c]
        diff=[]
        while i>0 or j>0:
            if(i>0 and j>0 and tb[i][j][0]=='u'):
                diff.append(['u',tb[i][0][0]])
                i-=1
                j-=1
            elif(j>0 and  tb[i][j][1]==tb[i][j-1][1]):
                diff.append(['d',tb[0][j][0]])
                j-=1
            else:
                diff.append(['i',tb[i][0][0]])
                i-=1
        return diff

    

    #两串比较
    @staticmethod
    def Compare(str1,str2):
        s=0
        len1=len(str1)
        e1=len(str1)-1
        e2=len(str2)-1

        #处理相同前导串
        while(s<=e1 and s<=e2 and str1[s]==str2[s]):
            s+=1
        #处理相同后辍串
        while(e1>=s and e2>=s and str1[e1]==str2[e2]):
            e1-=1
            e2-=1
        #比较中间差异s,e1,e2
        partialDiff=DiffLcs.genPartialDiff(str1,str2,s,e1,e2)

        #生成全部比对数组[{u|d|i},c]
        diff=[]
        for i in range(0,s):
            diff.append(['u',str1[i]])
        while(len(partialDiff)>0):
            diff.append(partialDiff.pop())
        for i in range(e1+1,len1):
            diff.append(['u',str1[i]])
        return diff
        
    #两文本文件比较    
    @staticmethod
    def CompareTextFile(file1,file2):
        fo1=open(file1)
        str1=fo1.read()
        fo2=open(file2)
        str2=fo2.read()
        return DiffLcs.Compare(str1,str2)

    #使用win32com调用word组件取doc内容文本
    #word为已初始化的VBAApplication
    @staticmethod
    def GetDocContent(word,file1):
        sdoc=word.Documents.Open(file1)
        str1= sdoc.Content.Text
        sdoc.Close()
        return str1
    
    #将diff比较结果写入doc文档
    @staticmethod
    def WriteDocDiff(word,file2,diff):
        ddoc=word.Documents.Open(file2)
        
        #清空文档内容
        i=1
        rng=ddoc.Range(0,len(ddoc.content.Text))
        rng.text=""
        
        #格式化比较结果 
        for t in diff:
            rng.InsertAfter(t[1])
            rng=ddoc.Range(i-1,i)
            if t[0]=='u':
                rng.font.StrikeThrough= False
                rng.font.color=0                
            if t[0]=='d':
                rng.font.StrikeThrough= True
                rng.font.color=255
            if t[0]=='i':
                rng.font.StrikeThrough= False
                rng.font.color=16711680
            i+=1
        #文首插入统计结果
        rng=ddoc.Range(0,0)
        rng.font.color=255
        rng.InsertAfter("[总字，正确，错录，漏录]\r\n")
        rng.InsertAfter(      DiffLcs.Count(diff))
        rng.InsertAfter("\r\n\r\n")
        ddoc.Save()
        ddoc.Close()

    #比较目录下所有doc文档
    #stdfile 为源标准doc文档
    #despath 为比较的目录
    @staticmethod
    def CompareDir(stdfile,despath):   
        #word.application组件初始化    
        word=win32com.client.Dispatch('Word.Application')
        word.Visible=0
        word.DisplayAlerts=0   

        #从doc读取取比对源串 
        stdStr=DiffLcs.GetDocContent(word,stdfile)

        word.Visible=1

        #取目录下的文件，循环处理
        for filename in os.listdir(despath):
            desFile=despath+ filename
            print(desFile)
            if os.path.isfile(desFile) and (desFile.endswith(".doc") or desFile.endswith(".docx")):



                desStr=DiffLcs.GetDocContent(word,desFile)    
                diff=DiffLcs.Compare(stdStr,desStr)
                print(desFile,DiffLcs.Count(diff))
                DiffLcs.WriteDocDiff(word,desFile,diff)

        word.Quit()

    
    @staticmethod
    def CompareDocFile(file1,file2,file3):
        word=win32com.client.Dispatch('Word.Application')
        word.Visible=0
        word.DisplayAlerts=0
        sdoc=word.Documents.Open(file1)
        str1= sdoc.Content.Text
        sdoc.Close()
        ddoc=word.Documents.Open(file2)
        str2=ddoc.Content.Text
        

        diff=DiffLcs.Compare(str1,str2)

        
        word.Visible=1
        i=1
        rng=ddoc.Range(0,len(ddoc.content.Text))
        rng.text=""
        
        for t in diff:
            rng.InsertAfter(t[1])
            rng=ddoc.Range(i-1,i)
            if t[0]=='u':
                rng.font.StrikeThrough= False
                rng.font.color=0                
            if t[0]=='d':
                rng.font.StrikeThrough= True
                rng.font.color=255
            if t[0]=='i':
                rng.font.StrikeThrough= False
                rng.font.color=16711680
            i+=1
        rng=ddoc.Range(0,0)
        rng.InsertAfter("[总字，正确，错录，漏录]\r\n")
        rng.InsertAfter(      DiffLcs.Count(diff))
        rng.InsertAfter("\r\n\r\n")


        #ddoc.Save()
        #cdoc.Close()
        #ddoc.Close()


        #word.Quit()
        return diff 
         
    

 
 


    @staticmethod
    def toHtml(diff):
        html=""
        for t in diff:
            if(t[0]=='u'):
                html+=t[1]
            elif(t[0]=='d'):
                html+="<font color='green'  style='text-decoration:underline'>"+ t[1]+"</font>"
            else:
                html+="<font color='red'>"+ t[1]+"</font>"     
        return html      

    @staticmethod
    def Count(diff):
        u=0
        d=0
        i=0
        c=len(diff)
        for t in diff:
            if(t[0]=='u'):
                u+=1
            elif(t[0]=='d'):
                d+=1
            else:
                i+=1
        result=[c,u,d,i]
        return result





 
#dp=DiffLcs.Compare(ss1,ss2)
#dp=DiffLcs.CompareTextFile("d:\\110-1192.TXT","d:\\110-11921.TXT")
#dp=DiffLcs.CompareDocFile("d:\\110-1192.doc","d:\\110-11921.doc","d:\\11-11922.doc")
#print(DiffLcs.toHtml(dp))
DiffLcs.CompareDir("d:\\110-1192.doc","d:\\test\\")
#print(DiffLcs.Count(dp))



 

 


