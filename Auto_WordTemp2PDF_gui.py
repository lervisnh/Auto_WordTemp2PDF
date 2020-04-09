# -*- coding:utf-8 -*-
# @Auther: lervisnh

from Auto_WordTemp2PDF_func import AUTO_WORDTemp2PDF
from tkinter import Frame, Button, Entry, StringVar, Checkbutton, IntVar
from tkinter.filedialog import askopenfilename, askdirectory
from tkinter.messagebox import showinfo
import os

class Application(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.pack_propagate(0)
        self.createWidgets()

    def createWidgets(self):
        Entry_width = 60

        fm1 = Frame(self.master)
        self.word_temp_Button = Button(fm1, text='选择Word模板', width=13,height=1, command=self.open_word_temp)
        self.word_temp_Button.pack(side='left')
        self.DocxTemplatePath = StringVar()
        self.DocxTemplatePath.set(u'请选择Word模板')
        self.word_temp_Entry= Entry(fm1, textvariable=self.DocxTemplatePath ,state='disabled',bd=0.5,width=Entry_width,)
        self.word_temp_Entry.pack(side='left')
        fm1.pack(side='top')

        fm2 = Frame(self.master)
        self.info_Button = Button(fm2, text='选择填入信息', width=13,height=1, command=self.open_infos)
        self.info_Button.pack(side='left')
        self.InfosPath = StringVar()
        self.InfosPath.set(u'请选择填入信息')
        self.info_Entry= Entry(fm2, textvariable=self.InfosPath, state='disabled',bd=0.5,width=Entry_width,)
        self.info_Entry.pack(side='left')
        fm2.pack(side='top')

        fm3 = Frame(self.master)
        self.saved_directory_Button = Button(fm3, text='选择保存路径', width=13,height=1, command=self.open_saved_directory)
        self.saved_directory_Button.pack(side='left')
        self.SaveDictionary = StringVar()
        self.SaveDictionary.set(u'请选择结果保存路径')
        self.saved_directory_Entry= Entry(fm3, textvariable=self.SaveDictionary, state='disabled',bd=0.5,width=Entry_width,)
        self.saved_directory_Entry.pack(side='left')
        fm3.pack(side='top')

        fm4 = Frame(self.master)
        fm4_1 = Frame(fm4)
        # 只保留PDF
        self.word_IntVar = IntVar()
        self.word_CheckButton = Checkbutton( fm4_1, text=u'只保留PDF', variable=self.word_IntVar )
        self.word_CheckButton.pack(side='left')
        fm4_1.pack(side='left')

        fm4_2 = Frame(fm4)
        # 开始生成
        self.start_working_Button = Button(fm4_2, text='开始生成', width=13,height=1, command=self.start_working)
        self.start_working_Button.pack(side='top')
        fm4_2.pack(side='left')

        fm4.pack(side='top')

    def open_word_temp(self,):
        self.DocxTemplatePath.set(askopenfilename(title=u'选择Word模板文件', 
                                                 filetypes=[('Word Files', '*.docx *.doc')], ))
        self.word_temp_Entry.select_clear()
        self.word_temp_Entry.insert( 0, self.DocxTemplatePath )
    
    def open_infos(self,):
        self.InfosPath.set(askopenfilename(title=u'选择填入信息文件', 
                                          filetypes=[('Excel files', "*.xlsx *.xls *.csv")],))
        self.info_Entry.select_clear()
        self.info_Entry.insert( 0, self.InfosPath )

    def open_saved_directory(self,):
        self.SaveDictionary.set(askdirectory(title=u'选择结果保存路径', ))
        self.saved_directory_Entry.select_clear()
        self.saved_directory_Entry.insert( 0, self.SaveDictionary )

    def start_working(self,):
        if self.DocxTemplatePath.get()==u'请选择Word模板' or not self.DocxTemplatePath.get():
            showinfo( title=u'请选择Word模板', message=u'请选择Word模板' )
            return 
        if self.InfosPath.get()==u'请选择填入信息' or not self.InfosPath.get():
            showinfo( title=u'请选择填入信息', message=u'请选择填入信息' )
            return 
        if self.SaveDictionary.get()==u'请选择结果保存路径' or not self.SaveDictionary.get():
            showinfo( title=u'请选择结果保存路径', message=u'请选择结果保存路径！' )
            return 
        
        try:
            obj = AUTO_WORDTemp2PDF(self.DocxTemplatePath.get(), self.InfosPath.get(), self.SaveDictionary.get() )
            obj.get_pdfs()
            self.delete_word()
            showinfo( title=u'自动生成成功', message=(u'自动生成成功，确定打开目录\n'+self.SaveDictionary.get()) )
            os.startfile(self.SaveDictionary.get())
        except:
            showinfo( title=u'自动生成失败', message=u'自动生成失败' )

    def delete_word(self,):
        if self.word_IntVar.get():
            print(u'开始删除Word文件...')
            for root, dirs, files in os.walk(self.SaveDictionary.get()):
                for file_name in files:
                    file_type = file_name.split('.')[-1]
                    if file_type=='docx' or file_type=='doc':
                        os.remove( os.path.join(root, file_name) )
                        print(os.path.join(root, file_name), u' 已删除')



if __name__=="__main__":
    app = Application()
    # 设置窗口标题:
    app.master.title('BY LERVISNH')
    # 主消息循环:
    app.mainloop()