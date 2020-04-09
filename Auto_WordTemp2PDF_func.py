# -*- coding:utf-8 -*-
# @Auther: lervisnh

from docxtpl import DocxTemplate
import pandas as pd
from datetime import date
import os
from win32com.client import Dispatch

class AUTO_WORDTemp2PDF:
    def __init__(self, DocxTemplatePath, InfosPath, SaveDictionary):
        # DocxTemplatePath 为word模板地址，字符串类型
        # InfosPath 为填入信息表格地址，可为xls/xlsx/csv
        # SaveDictionary 保存文件的路径，是文件夹
        self.DocxTemplatePath = DocxTemplatePath
        self.InfosPath = InfosPath
        self.SaveDictionaryPath = SaveDictionary

    def __ReadTemplate(self,):
        return DocxTemplate(self.DocxTemplatePath)
    
    def __ReadInfos(self,):
        if self.InfosPath.split('.')[-1]=='csv':
            self.Infos = pd.read_csv(self.InfosPath)
        else:
            self.Infos = pd.read_excel(self.InfosPath)
    
    def __CheckSaveDictionary(self,):
        if not os.path.exists(self.SaveDictionaryPath):
            os.makedirs( self.SaveDictionaryPath )

    def __FillContents(self,):
        # 检查填入信息是否加载
        try:
            hasattr( self, Infos )
        except:
            self.__ReadInfos()
        # 检查保存路径是否存在
        self.__CheckSaveDictionary()
        WordFiles = []
        # 开始填写
        for idx, row in self.Infos.iterrows():
            myDocxTemplate = self.__ReadTemplate()#加载模板
            contents = row.to_dict()
            #print(contents)
            contents['y_m_d'] = str(date.today())
            myDocxTemplate.render( contents )
            #保存填写后的word
            if 'phone' in contents:
                WordFile = self.SaveDictionaryPath+'/'+ contents['name'] + str(contents['phone']) +'.docx'
            else:
                WordFile = self.SaveDictionaryPath+'/'+ contents['name'] +'.docx'
            # 生成文件
            if not os.path.isfile( WordFile ):
                myDocxTemplate.save(WordFile)
                print( WordFile+u' 生成成功' )
            else:
                print( WordFile+u' 已存在' )
            WordFiles.append(WordFile)
        # 生成pdf
        self.__Word2PDF( WordFiles )
            
    def __Word2PDF(self, wordfilePaths):
        # 调用office word 或者 wps
        try:
            w = Dispatch("Word.Application")
        except:
            w = Dispatch("wps.Application")
        # 开始转PDF
        print(u'\n开始生成PDF...')
        try:
            for wordfilePath in wordfilePaths:
                doc = w.Documents.Open(wordfilePath, ReadOnly=1)
                pdfPath = wordfilePath.replace(".docx", ".pdf")
                if not os.path.isfile(pdfPath):
                    doc.ExportAsFixedFormat(pdfPath, 17, Item=7, CreateBookmarks=1)
                    print( wordfilePath.replace(".docx", ".pdf") + u' 生成成功' )
                else:
                    print( pdfPath+u'已存在' )
        except Exception as error:
            print( 'Please Check Error: ' )
            print( error )
        finally:
            doc.Close() # 关闭文档
            w.Quit(0)

    def get_pdfs(self,):
        self.__FillContents()


if __name__=='__main__':
    '''
    DocxTemplatePath = 'D:/Projects/python/Auto_WordTemp2PDF/test_case/奖状模板.docx'
    InfosPath = 'D:/Projects/python/Auto_WordTemp2PDF/test_case/学生信息.xlsx'
    SaveDictionary = 'D:/Projects/python/Auto_WordTemp2PDF/test_case/outputs'
    '''
    DocxTemplatePath = input(u'\n\nWORD模板地址(绝对路径)：\n')
    while not os.path.isfile(DocxTemplatePath):
        DocxTemplatePath = input(u'WORD模板地址错误，请重新输入：')
    InfosPath = input(u'填入信息地址(绝对路径)：\n')
    while not os.path.isfile(InfosPath):
        InfosPath = input(u'填入信息地址错误，请重新输入：')
    SaveDictionary = input(u'保存目录(绝对路径)：\n')
    
    obj = AUTO_WORDTemp2PDF(DocxTemplatePath, InfosPath, SaveDictionary)

    obj.get_pdfs()
    
    input('\n已完成，回车退出！')