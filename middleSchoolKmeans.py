# -*- coding:utf-8 -*-
import time
import xlrd
import xlwt
from numpy import *

def loadAndDealData():
    workbook = xlrd.open_workbook('math2.xls')
    dealsheet = workbook.sheet_by_index(0)
    scoreList =[]
    for i in range(1,dealsheet.nrows):
        if dealsheet.cell_value(i,4) != '':
            scoreList.append(dealsheet.cell_value(i,4))
    return scoreList

def writeData(curList,mycenter,k):
    start =65;startRank = curList[0]
    workbook = xlrd.open_workbook('math2.xls')
    newexcel = xlwt.Workbook()
    newSheet = newexcel.add_sheet(u'sheet1',cell_overwrite_ok=True)
    dealsheet = workbook.sheet_by_index(0)
    sign1 = 0
    for i in range(1, dealsheet.nrows):
        if dealsheet.cell_value(i, 4) != '':
            sign1 +=1
            for j in range(0,dealsheet.ncols):
                newSheet.write(i,j,dealsheet.cell_value(i,j))
            newSheet.write(i,dealsheet.ncols,curList[i-1])
            if startRank == curList[i-1]:
                newSheet.write(i,dealsheet.ncols+1,chr(start))
            else:
                startRank = curList[i-1]
                start+=1
                newSheet.write(i,dealsheet.ncols+1,chr(start))

    newSheet.write(0, 0, '分析结果')
    newSheet.write(0, 1, '四类的中心点')
    newSheet.write(0, 7, '聚类的评价等级')
    for item in range(k):
        newSheet.write(0,2+item,float(mycenter[item,][0]))
    newexcel.save('newdata.xls')



def distEclud(score1,score2):
    return sqrt(power(score2-score1,2))

def randCent(datalist,k):
    centroids = mat(zeros((k,1)))
    minScore = min(datalist)
    rangeScore = float(max(datalist)-minScore)
    centroids[:,0] = minScore + rangeScore*random.rand(k,1)
    return centroids

def scoreKmeans(datalist,k,dist=distEclud,creatCenter=randCent):
    flag = 0
    prescoreMat = mat(datalist)
    scoreMat = prescoreMat.reshape((len(datalist), 1))
    num = len(datalist)
    clusterArray = mat(zeros((num,2)))
    centroids = creatCenter(datalist,k)
    clusterChanged=True
    while clusterChanged:
        flag +=1
        clusterChanged = False
        for i in range(num):
            minDist = inf
            minIndex = -1
            for j in range(k):
                distNow = dist(centroids[j,:],datalist[i])
                if distNow < minDist:
                    minDist = distNow
                    minIndex = j
            if clusterArray[i,0] != minIndex:
                clusterChanged =True
            clusterArray[i,:]=minIndex,minDist
        #print(clusterArray)
        for center in range(k):
            ptsCluter = scoreMat[nonzero(clusterArray[:,0].A==center)[0]]
            #print(nonzero(clusterArray[:,0].A==center)[0])
            centroids[center,:]=mean(ptsCluter,axis=0)
        result = ''
        with open('result.txt','a') as _file:
            _file.write(u'第%d次迭代' %flag)
            for item in range(centroids.shape[0]):
                result +=str(float(centroids[item,][0]))+' '
            _file.write(result)
            _file.write('\n'+'*********************'+'\n\n\n')

    return centroids,clusterArray,flag

if __name__ == '__main__':
    k=4
    datalist = loadAndDealData()
    with open('result.txt', 'a') as _file:
        _file.write(u'log' + str(time.time()) + '\n')
    mycenter,cluserArray,flag = scoreKmeans(datalist,k)
    print (u'处理完成，迭代了%d次' %flag)
    clusterList = []
    # for i in range(mycenter.shape[0]):
    #     clusterList.append(float(mycenter[i,][0]))
    # clusterList.sort(reverse=True)
    #writeData(cluserArray,mycenter)
    #print(cluserArray)
    # print(array(cluserArray[0][0]).shape)
    # newarray=array(cluserArray[0][0])
    # print(newarray[0][0])
    for i in range(cluserArray.shape[0]):
         newarray = array(cluserArray[i][0])
         clusterList.append(int(newarray[0][0]))
    writeData(clusterList,mycenter,k)