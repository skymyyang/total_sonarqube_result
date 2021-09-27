#!/usr/bin/env python
# coding=utf-8


import xlwings as xw
import time


class write_table_row():
    def __init__(self, data):
        self.data = data

    # def create_table(slef):
    #     now = time.strftime(r"%Y-%m-%d")
    #     app = xw.App(visible=False, add_book=False)
    #     app.display_alerts=True
    #     app.screen_updating=False
    #     wb = app.books.add()
    #     wb.sheets["sheet1"].range('A1').value = ["项目名称", "BUGS", "漏洞", "异味"]
    #     wb.save('sonarqube'+ now + '.xls')
    #     wb.close()
    #     app.quit()
    #     return 'sonarqube'+ now + '.xls'
    

    # def write_data(self, xlsfile):
    #     app = xw.App(visible=False, add_book=False)
    #     app.display_alerts=True
    #     app.screen_updating=False
    #     wb = app.books.open(xlsfile)
    #     wb.sheets['sheet1'].range('A2').value = self.data
    #     wb.save()
    #     wb.close()
    #     app.quit()

    def wirte_data(self):
        """定义好要存入列表中数据的类型"""
        now = time.strftime(r"%Y-%m-%d")
        app = xw.App(visible=False, add_book=False)
        app.display_alerts=True
        app.screen_updating=False
        wb = app.books.add()
        sht = wb.sheets.active
        
        sht.range('A1').value = ["项目名称", "BUGS", "漏洞", "异味", "代码行数", "严重程度[BUG]阻断-严重-主要-次要-提示", "严重程度[漏洞]阻断-严重-主要-次要-提示", "严重程度[异味]阻断-严重-主要-次要-提示"]
        sht.range('A2').value = self.data
        sht.autofit()
        #wb.sheets["sheet1"].range('A1').value = ["项目名称", "BUGS", "漏洞", "异味", "代码行数", "严重程度[BUG]阻断-严重-主要-次要-提示", "严重程度[漏洞]阻断-严重-主要-次要-提示", "严重程度[漏洞]阻断-严重-主要-次要-提示"]
        #wb.sheets["sheet1"].range('A2').value = self.data
        wb.save('sonarqube'+ now + '.xls')
        wb.close()
        app.quit()




# s = write_table_row("123")
# s.create_table()

#nowdate = datetime.datetime.now()
#nowdate_s = str(nowdate.year)+"-"+str(nowdate.month)
#+nowdate.month+nowdate.day
#print(nowdate_s)
# now_name = datetime.datetime.strptime(nowdate, '%Y-%m-%d')
# print(now_name)

# now = time.strftime(r"%Y-%m-%d")
# print(now)

