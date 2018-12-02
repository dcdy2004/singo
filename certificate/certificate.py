import xlrd
from docxtpl import DocxTemplate
import random
import os
class certificate():
    file = './info.xls'
    tpl = 0
    content=[]
    def read_info(self):
        workbook = xlrd.open_workbook(self.file)
        sheet = workbook.sheet_by_index(0)
        for i in range(0,sheet.nrows):

            if(sheet.row_values(i)[0]==i):
                rows = sheet.row_values(i)
                self.content.append(rows)
    def generate_err(self):
        err=[[],[]]
        err[0].append(0.2+random.random()*0.15)
        err[0].append(err[0][0] /(random.random(*0.1)+1.3))
        err[0].append(err[0][1] /(random.random(*0.1)+1.4))
        err[0].append(err[0][2] /(random.random(*0.1)+1.1))
        err[0].append(err[0][3] /(random.random(*0.1)+1))
        err[1].append(10+random.random()*5)
        err[1].append(err[0][0]/(random.random(*0.1)+1.4))
        err[1].append(err[0][1]/(random.random(*0.1)+1.5))
        err[1].append(err[0][2]/(random.random(*0.1)+1.9))
        err[1].append(err[0][3]/(random.random(*0.1)+1.4))
        return err

    def generate(self):
        err=[[],[]]
        print(len(self.content))
        for i in range(len(self.content)):
            if self.content[i][2]=='电流互感器':
                if self.content[i][3]==10:
                    self.tpl=DocxTemplate('./ct10.docx')
                elif self.content[i][2]==35:
                    self.tpl=DocxTemplate('./ct35.docx')
                elif self.content[i][2]==110:
                    self.tpl = DocxTemplate('./ct110.docx')
            elif self.content[i][1]=='电压互感器':
                if self.content[i][2]==10:
                    self.tpl = DocxTemplate('./pt10.docx')
                elif self.content[i][2]==35:
                    self.tpl=DocxTemplate('./pt35.docx')
                elif self.content[i][2]==110:
                    self.tpl = DocxTemplate('./pt110.docx')
            company=self.content[i][1]
            if(len(self.content[i][1])<19):
                company='_'*int((24-len(self.content[i][1]))/2)+self.content[i][1]
                company+='_'*(24-len(company))

            self.tpl.render({
                'company':company

            })
            self.tpl.save('./dddd .docx')

if __name__ =='__main__':
    new = certificate()
    new.read_info()
    new.generate()
