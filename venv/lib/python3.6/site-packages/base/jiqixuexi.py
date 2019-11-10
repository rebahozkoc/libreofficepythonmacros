#coding:utf8
import os
import os.path
import codecs
import numpy
import pandas
import jieba
import base

#继续Base
class Base(base.Base):


    #余弦值计算函数( 传入2个向量)
    def cosineDist(col1, col2):
        return numpy.sum(col1 * col2)/(
            numpy.sqrt(numpy.sum(numpy.power(col1, 2))) *
            numpy.sqrt(numpy.sum(numpy.power(col2, 2)))
        )

    #有的时候需要用到这个列表可以用一下
    def get_stopwords(self):
        #停用词列表
        return pandas.read_csv(
            "StopwordsCN.txt",
            encoding='utf8',
            index_col=False,
            quoting=3,
            sep="\t"
        )

    #默认是去除停用词,如果设置0,就不去除 -----
    #针对只要的算法[文章,句子,句子,句子],如果仅仅要切分内容,就直接传入[内容],得出的就是一个分词以后内容了.
    def fenci(self,subCorpos,cutstopword=1):
        #分词且去除停用词
        sub = []
        segments = []
        for j in range(len(subCorpos)):
            segs = jieba.cut(subCorpos[j])
            for seg in segs:
                if len(seg.strip())>1:
                    sub.append(subCorpos[j])    #目标段落 & 目标文章
                    segments.append(seg)        #对应分词
        #不去除停用词
        segmentDF = pandas.DataFrame({'sub':sub, 'segment':segments})
        if cutstopword == 0:
            return segmentDF

        #返回去除停用词
        else:
            return segmentDF[~segmentDF.segment.isin(stopwords.stopword)]



