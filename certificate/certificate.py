import xlrd
from docxtpl import DocxTemplate
import random
import datetime
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
        err[0].append(0.05+random.random()*0.2)
        err[0].append(err[0][0] /(random.random()*0.1+1.3))
        err[0].append(err[0][1] /(random.random()*0.2+1.8))
        err[0].append(err[0][2] /(random.random()*0.3+1.7))
        err[0].append(err[0][3] /(random.random()*0.1+1.1))
        err[1].append(10+random.random()*5)
        err[1].append(err[1][0]/(random.random()*0.3+1.5))
        err[1].append(err[1][1]/(random.random()*0.4+1.1))
        err[1].append(err[1][2]/(random.random()*0.5+1.4))
        err[1].append(err[1][3]/(random.random()*0.4+1.1))
        return err

    def generate(self):
        err=[[],[]]
        for i in range(len(self.content)):
            if self.content[i][2]=='电流互感器':
                if self.content[i][3]=='10kV':
                    self.tpl=DocxTemplate('./ct10.docx')
                elif self.content[i][2]=='35kV':
                    self.tpl=DocxTemplate('./ct35.docx')
                elif self.content[i][2]=='110kV':
                    self.tpl = DocxTemplate('./ct110.docx')
            elif self.content[i][1]=='电压互感器':
                if self.content[i][2]=='10kV':
                    self.tpl = DocxTemplate('./pt10.docx')
                elif self.content[i][2]=='35KV':
                    self.tpl=DocxTemplate('./pt35.docx')
                elif self.content[i][2]=='110kV':
                    self.tpl = DocxTemplate('./pt110.docx')
#company
            company=self.content[i][1]
            if(len(self.content[i][1])<19):
                company='_'*int((40-len(self.content[i][1])*2)/2)+self.content[i][1]  #40减掉两倍中文字符长度的一半是前面'_'的数量
                company+='_'*(40-int((40-len(self.content[i][1])*2)/2)-len(self.content[i][1])*2)
            else :
                company = '名称过长，请手动修改'
#model
            model='_'*int((40-len(self.content[i][4]))/2)+self.content[i][4]+'_'*(40-int((40-len(self.content[i][4]))/2)-len(self.content[i][4]))
#serial
            if (len(str(self.content[i][7]))<1):
                if(len(str(int(self.content[i][5])))<=19):
                    serial = str(int(self.content[i][5])) + '  ' + str(int(self.content[i][6]))
                    serial = '_' * int((40 - len(serial)) / 2) + serial + '_' * (40 - int((40 - len(serial)) / 2) - len(serial))
                else:
                    serial='编号过长，请手动修改'
            else:
                if (len(str(int(self.content[i][5]))) <= 12):
                    serial = str(int(self.content[i][5])) + '  ' + str(int(self.content[i][6]))+'  ' + str(int(self.content[i][7]))
                    serial = '_' * int((40 - len(serial)) / 2) + serial + '_' * (40 - int((40 - len(serial)) / 2) - len(serial))
                else:
                    serial = '编号过长，请手动修改'
#make
            make = self.content[i][8]
            if (len(self.content[i][8]) < 19):
                make = '_' * int((40 - len(self.content[i][8]) * 2) / 2) + self.content[i][8]  # 40减掉两倍中文字符长度的一半是前面'_'的数量
                make += '_' * (40 - int((40 - len(self.content[i][8]) * 2) / 2) - len(self.content[i][8]) * 2)
            else:
                make = '名称过长，请手动修改'
#date
            date_tmp=datetime.datetime.strftime(datetime.datetime.now(),'%Y%m%d')
            if(date_tmp[-2:]=='01'):
                date_tmp=date_tmp[:-2]+'02'
            date='     '+date_tmp[:4]+'     年'+'   '+date_tmp[4:6]+'   月'+'   '+date_tmp[-2:]+'   日'
            date2='     '+str(int(date_tmp[:4])+10)+'     年'+'   '+date_tmp[4:6]+'   月'+'   '
            if(int(date_tmp[-2:])<10):
                date2+='0'
            date2+=str(int(date_tmp[-2:])-1)+'   日'
#makedate
            makedate = '_' * int((40 - len(str(self.content[i][9]))) / 2) + str(self.content[i][9]) + '_' * (40 - int((40 - len(str(self.content[i][9]))) / 2) - len(str(self.content[i][9])))
#primary
            primary = '_' * int((39 - len(str(self.content[i][10]))) / 2) + str(self.content[i][10])[:-2] + '_' * (39 - int((39 - len(str(self.content[i][10]))) / 2) - len(str(self.content[i][10])))
#load
            load='_' * int((39 - len(str(self.content[i][11]))) / 2) + str(self.content[i][11])[:-2] + '_' * (39 - int((39 - len(str(self.content[i][11]))) / 2) - len(str(self.content[i][11])))
            load2=str(self.content[i][11])[:-2]
#standard
            standerd = '_' * int((37 - len(str(self.content[i][3][:-2]))) / 2) + str(self.content[i][3][:-2]) + '_' * (38 - int((38 - len(str(self.content[i][3][:-2]))) / 2) - len(str(self.content[i][3][:-2])))
#err
            err1=self.generate_err()
            err2=self.generate_err()
            err3=self.generate_err()
            err4=self.generate_err()
            err11 = '+'+str(err1[0][0])[:5]
            err12 = '+'+str(err1[0][1])[:5]
            err13 = '+'+str(err1[0][2])[:5]
            err14 = '+'+str(err1[0][3])[:5]
            err15 = '+'+str(err1[0][4])[:5]
            err16 = '+'+str(err2[0][0])[:5]
            err17 = '+'+str(err2[0][1])[:5]
            err18 = '+'+str(err2[0][2])[:5]
            err19 = '+'+str(err2[0][3])[:5]
            err21 = '+'+str(err1[1][0])[:5]
            err22 = '+'+str(err1[1][1])[:4]
            err23 = '+'+str(err1[1][2])[:4]
            err24 = '+'+str(err1[1][3])[:4]
            err25 = '+'+str(err1[1][4])[:4]
            err26 = '+'+str(err2[1][0])[:5]
            err27 = '+'+str(err2[1][1])[:4]
            err28 = '+'+str(err2[1][2])[:4]
            err29 = '+'+str(err2[1][3])[:4]
            err31 = '+'+str(err3[0][0])[:5]
            err32 = '+'+str(err3[0][1])[:5]
            err33 = '+'+str(err3[0][2])[:5]
            err34 = '+'+str(err3[0][3])[:5]
            err35 = '+'+str(err3[0][4])[:5]
            err36 = '+'+str(err4[0][0])[:5]
            err37 = '+'+str(err4[0][1])[:5]
            err38 = '+'+str(err4[0][2])[:5]
            err39 = '+'+str(err4[0][3])[:5]
            err41 = '+'+str(err3[1][0])[:5]
            err42 = '+'+str(err3[1][1])[:4]
            err43 = '+'+str(err3[1][2])[:4]
            err44 = '+'+str(err3[1][3])[:4]
            err45 = '+'+str(err3[1][4])[:4]
            err46 = '+'+str(err4[1][0])[:5]
            err47 = '+'+str(err4[1][1])[:4]
            err48 = '+'+str(err4[1][2])[:4]
            err49 = '+'+str(err4[1][3])[:4]
#equip3
            if (len(str(self.content[i][7])) < 1):
                load3='/'
                lite='/'
                err51='/'
                err52='/'
                err53='/'
                err54='/'
                err55='/'
                err56='/'
                err57='/'
                err58='/'
                err59='/'
                err61='/'
                err62='/'
                err63='/'
                err64='/'
                err65='/'
                err66='/'
                err67='/'
                err68='/'
                err69='/'
                serial3=''
            else:
                err5 = self.generate_err()
                err6 = self.generate_err()
                load3 = load2
                lite = '3.75'
                err51 = '+' + str(err5[0][0])[:5]
                err52 = '+' + str(err5[0][1])[:5]
                err53 = '+' + str(err5[0][2])[:5]
                err54 = '+' + str(err5[0][3])[:5]
                err55 = '+' + str(err5[0][4])[:5]
                err56 = '+' + str(err6[0][0])[:5]
                err57 = '+' + str(err6[0][1])[:5]
                err58 = '+' + str(err6[0][2])[:5]
                err59 = '+' + str(err6[0][3])[:5]
                err61 = '+' + str(err5[1][0])[:5]
                err62 = '+' + str(err5[1][1])[:4]
                err63 = '+' + str(err5[1][2])[:4]
                err64 = '+' + str(err5[1][3])[:4]
                err65 = '+' + str(err5[1][4])[:4]
                err66 = '+' + str(err6[1][0])[:5]
                err67 = '+' + str(err6[1][1])[:4]
                err68 = '+' + str(err6[1][2])[:4]
                err69 = '+' + str(err6[1][3])[:4]
            self.tpl.render({
                'company':company,
                'model':model,
                'serial':serial,
                'make':make,
                'date':date,
                'date2':date2,
                'makedate':makedate,
                'primary':primary,
                'load':load,
                'load2':load2,
                'load3':load3,
                'lite':lite,
                'standard':standerd,
                'serial1':str(int(self.content[i][5])),
                'serial2':str(int(self.content[i][6])),
                'serial3':serial3,
                'err11' :err11,
                'err12' :err12,
                'err13' :err13,
                'err14' :err14,
                'err15' :err15,
                'err16' :err16,
                'err17' :err17,
                'err18' :err18,
                'err19' :err19,
                'err21' :err21,
                'err22' :err22,
                'err23' :err23,
                'err24' :err24,
                'err25' :err25,
                'err26' :err26,
                'err27' :err27,
                'err28' :err28,
                'err29' :err29,
                'err31' :err31,
                'err32' :err32,
                'err33' :err33,
                'err34' :err34,
                'err35' :err35,
                'err36' :err36,
                'err37' :err37,
                'err38' :err38,
                'err39' :err39,
                'err41' :err41,
                'err42' :err42,
                'err43' :err43,
                'err44' :err44,
                'err45' :err45,
                'err46' :err46,
                'err47' :err47,
                'err48' :err48,
                'err49' :err49,
                'err51': err51,
                'err52': err52,
                'err53': err53,
                'err54': err54,
                'err55': err55,
                'err56': err56,
                'err57': err57,
                'err58': err58,
                'err59': err59,
                'err61': err61,
                'err62': err62,
                'err63': err63,
                'err64': err64,
                'err65': err65,
                'err66': err66,
                'err67': err67,
                'err68': err68,
                'err69': err69
            })
            j=1
            if(self.content[i][2]=='电流互感器'):
                type='CT'
            else:
                type='PT'
            while(os.path.exists(self.content[i][1]+type+str(j)+'.docx')==True):
                j+=1
            self.tpl.save(self.content[i][1]+type+str(j)+'.docx')

if __name__ =='__main__':
    new = certificate()
    new.read_info()
    new.generate()
