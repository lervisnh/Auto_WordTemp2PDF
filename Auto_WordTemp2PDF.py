# -*- coding:utf-8 -*-
# @Auther: lervisnh

from Auto_WordTemp2PDF_gui import Application

if __name__=="__main__":
    app = Application()
    # 设置窗口标题:
    app.master.title('Word模板批量生成PDF BY LERVISNH')
    # 居中显示
    app.master.update() # update window ,must do
    curWidth = app.master.winfo_reqwidth() # get current width
    curHeight = app.master.winfo_height() # get current height
    scnWidth,scnHeight = app.master.maxsize() # get screen width and height
    # now generate configuration information
    tmpcnf = '%dx%d+%d+%d'%(curWidth,curHeight, (scnWidth-curWidth)/2,(scnHeight-curHeight)/2)
    app.master.geometry(tmpcnf)
    # 锁定窗口大小
    app.master.resizable(0,0)
    # 主消息循环:
    app.mainloop()