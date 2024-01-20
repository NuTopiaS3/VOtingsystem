from kivymd.app import MDApp
from kivy.lang import Builder
from kivy.uix.screenmanager import Screen,ScreenManager
from kivy.animation import Animation
import pandas
from openpyxl import Workbook
import openpyxl
import os,sys
from kivy.resources import resource_add_path, resource_find


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class LoginScreen(Screen):
    pass
class PrefectScreen(Screen):
    pass
class PrefectgScreen(Screen):
    pass
class StudentScreen(Screen):
    pass
class Maple(Screen):
    pass
class Cedar(Screen):
    pass
class Oak(Screen):
    pass
class Pinec(Screen):
    pass
class Pinev(Screen):
    pass


kv='''
<LoginScreen>:
    name: 'log'
    canvas:
        Rectangle:
            source: resource_path('D:/Gugan/VOting system/bg.jpg')
            size: self.size
    MDCard:
        id:bg
        md_bg_color: "white"
        size_hint_x:0.336
        size_hint_y:0.6
        pos_hint: {'center_x': 0.5 ,'center_y': 0.5}

    MDTextField:
        id: user
        hint_text: 'User Name'
        size_hint_x: 0.3
        pos_hint:{'center_x': 0.5 ,'center_y': 0.6}
        line_color_normal: 0,205/255,1,1
        hint_text_color_normal: 0,205/255,1,1

    MDTextField:
        id: passw
        hint_text: 'Password'
        helper_text: 'Your Adm No. '
        size_hint_x: 0.3
        pos_hint:{'center_x': 0.5 ,'center_y': 0.48}
        line_color_normal: 0,205/255,1,1
        hint_text_color_normal: 0,205/255,1,1
        password: True

    MDRoundFlatButton:
        text: 'Login'
        pos_hint:{'center_x': 0.5 ,'center_y': 0.33}
        size_hint_x: 0.15
        on_release: app.login(user, passw,error)

    Label:
        id: error
        text: ''
        color: 1,0,0,1
        pos_hint:{'center_x': 0.5 ,'center_y': 0.27}

        
<PrefectgScreen>:
    name: 'votepg'
    canvas:
        Rectangle:
            source: resource_path('D:/Gugan/VOting system/bg.jpg')
            size: self.size 

    Label:
        text: 'Girl Prefect'
        pos_hint: {'center_x':0.5,'center_y':0.88}
        font_size: 72
        bold: True
    Label:
        text:'NAME'
        pos_hint: {'center_x':0.0875,'center_y':0.38}
        color:0,0,1,1

    Label:
        text:'NAME'
        pos_hint: {'center_x':0.3625,'center_y':0.38}
        color:0,0,1,1

    Label:
        text:'NAME'
        pos_hint: {'center_x':0.6375,'center_y':0.38}
        color:0,0,1,1

    Label:
        text:'NAME'
        pos_hint: {'center_x':0.9125,'center_y':0.38}
        color:0,0,1,1        

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.175
        size_hint_y:0.17
        pos_hint: {'center_x':0.0875,'center_y':0.49}

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.175
        size_hint_y:0.17
        pos_hint: {'center_x':0.3625,'center_y':0.49}

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.175
        size_hint_y:0.17
        pos_hint: {'center_x':0.6375,'center_y':0.49}

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.175
        size_hint_y:0.17
        pos_hint: {'center_x':0.9125,'center_y':0.49}

    MDRaisedButton:
        id: but
        text: 'VOTE'
        pos_hint: {'center_x':0.0875,'center_y':0.331}
        size_hint_x:0.175
        opacity: 1  
        on_press: app.pgcontv1()
        disabled: False

    MDRaisedButton:
        id: but1
        text: 'VOTE'
        pos_hint: {'center_x':0.3625,'center_y':0.331}
        size_hint_x:0.175
        opacity: 1 
        on_press: app.pgcontv2()
        disabled: False

    MDRaisedButton:
        id: but2
        text: 'VOTE'
        pos_hint: {'center_x':0.6375,'center_y':0.331}
        size_hint_x:0.175
        opacity: 1
        on_press: app.pgcontv3()
        disabled: False

    MDRaisedButton:
        id: but2
        text: 'VOTE'
        pos_hint: {'center_x':0.9125,'center_y':0.331}
        size_hint_x:0.175
        opacity: 1
        on_press: app.pgcontv4()
        disabled: False   
        
<PrefectScreen>:
    name: 'votepb'
    canvas:
        Rectangle:
            source: resource_path('D:/Gugan/VOting system/bg.jpg')
            size: self.size 

    Label:
        text: 'Boy Prefect'
        pos_hint: {'center_x':0.5,'center_y':0.88}
        font_size: 72
        bold: True
    Label:
        text:'NAME'
        pos_hint: {'center_x':0.18,'center_y':0.38}
        color:0,0,1,1

    Label:
        text:'NAME'
        pos_hint: {'center_x':0.5,'center_y':0.38}
        color:0,0,1,1

    Label:
        text:'NAME'
        pos_hint: {'center_x':0.82,'center_y':0.38}
        color:0,0,1,1

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.18,'center_y':0.49}

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.5,'center_y':0.49}

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.82,'center_y':0.49}

    MDRaisedButton:
        id: but
        text: 'VOTE'
        pos_hint: {'center_x':0.18,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1  
        on_press: app.pbcontv1()
        disabled: False

    MDRaisedButton:
        id: but1
        text: 'VOTE'
        pos_hint: {'center_x':0.5,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1 
        on_press: app.pbcontv2()
        disabled: False

    MDRaisedButton:
        id: but2
        text: 'VOTE'
        pos_hint: {'center_x':0.82,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1
        on_press: app.pbcontv3()
        disabled: False   

<StudentScreen>:
    name: 'votes'
    canvas:
        Rectangle:
            source: resource_path('D:/Gugan/VOting system/bg.jpg')
            size: self.size 

            
    Label:
        text: 'Student Council'
        pos_hint: {'center_x':0.5,'center_y':0.88}
        font_size: 72
        bold: True

    Label:
        text:'NAME'
        pos_hint: {'center_x':0.18,'center_y':0.38}
        color:0,0,1,1

    Label:
        text:'NAME'
        pos_hint: {'center_x':0.5,'center_y':0.38}
        color:0,0,1,1

    Label:
        text:'NAME'
        pos_hint: {'center_x':0.82,'center_y':0.38}
        color:0,0,1,1

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.18,'center_y':0.49}

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.5,'center_y':0.49}

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.82,'center_y':0.49}

    MDRaisedButton:
        id: but
        text: 'VOTE'
        pos_hint: {'center_x':0.18,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1  
        on_press: app.scontv1()
        disabled: False

    MDRaisedButton:
        id: but1
        text: 'VOTE'
        pos_hint: {'center_x':0.5,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1 
        on_press: app.scontv2()
        disabled: False

    MDRaisedButton:
        id: but2
        text: 'VOTE'
        pos_hint: {'center_x':0.82,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1
        on_press: app.scontv3()
        disabled: False

<Maple>:
    name: 'maple'
    canvas:
        Rectangle:
            source: resource_path('D:/Gugan/VOting system/bg.jpg')
            size: self.size 

    Label:
        text: 'Maple Vice Captain'
        pos_hint: {'center_x':0.5,'center_y':0.88}
        font_size: 72
        bold: True

    Label:
        text:'NAME'
        pos_hint: {'center_x':0.18,'center_y':0.38}
        color:0,0,1,1

    Label:
        text:'NAME'
        pos_hint: {'center_x':0.5,'center_y':0.38}
        color:0,0,1,1

    Label:
        text:'NAME'
        pos_hint: {'center_x':0.82,'center_y':0.38}
        color:0,0,1,1

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.18,'center_y':0.49}

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.5,'center_y':0.49}

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.82,'center_y':0.49}

    MDRaisedButton:
        id: but
        text: 'VOTE'
        pos_hint: {'center_x':0.18,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1  
        on_press: app.mcontv1()
        disabled: False

    MDRaisedButton:
        id: but1
        text: 'VOTE'
        pos_hint: {'center_x':0.5,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1 
        on_press: app.mcontv2()
        disabled: False

    MDRaisedButton:
        id: but2
        text: 'VOTE'
        pos_hint: {'center_x':0.82,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1
        on_press: app.mcontv3()
        disabled: False

<Cedar>:
    name: 'cedar'
    canvas:
        Rectangle:
            source: resource_path('D:/Gugan/VOting system/bg.jpg')
            size: self.size 

    Label:
        text: 'Cedar Vice Captain'
        pos_hint: {'center_x':0.5,'center_y':0.88}
        font_size: 72
        bold: True

    Label:
        text:'NAME'
        pos_hint: {'center_x':0.18,'center_y':0.38}
        color:0,0,1,1

    Label:
        text:'NAME'
        pos_hint: {'center_x':0.5,'center_y':0.38}
        color:0,0,1,1

    Label:
        text:'NAME'
        pos_hint: {'center_x':0.82,'center_y':0.38}
        color:0,0,1,1

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.18,'center_y':0.49}

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.5,'center_y':0.49}

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.82,'center_y':0.49}

    MDRaisedButton:
        id: but
        text: 'VOTE'
        pos_hint: {'center_x':0.18,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1  
        on_press: app.ccontv1()
        disabled: False

    MDRaisedButton:
        id: but1
        text: 'VOTE'
        pos_hint: {'center_x':0.5,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1 
        on_press: app.ccontv2()
        disabled: False

    MDRaisedButton:
        id: but2
        text: 'VOTE'
        pos_hint: {'center_x':0.82,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1
        on_press: app.ccontv3()
        disabled: False

<Oak>:
    name: 'oak'
    canvas:
        Rectangle:
            source: resource_path('D:/Gugan/VOting system/bg.jpg')
            size: self.size 

    Label:
        text: 'Oak Captain'
        pos_hint: {'center_x':0.5,'center_y':0.88}
        font_size: 72
        bold: True

    Label:
        text:'NAME'
        pos_hint: {'center_x':0.3033,'center_y':0.38}
        color:0,0,1,1


    Label:
        text:'NAME'
        pos_hint: {'center_x':0.6075,'center_y':0.38}
        color:0,0,1,1

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.3033,'center_y':0.49}


    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.6075,'center_y':0.49}

    MDRaisedButton:
        id: but
        text: 'VOTE'
        pos_hint: {'center_x':0.3033,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1  
        on_press: app.ocontv1()
        disabled: False


    MDRaisedButton:
        id: but2
        text: 'VOTE'
        pos_hint: {'center_x':0.6075,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1
        on_press: app.ocontv3()
        disabled: False

<Pinec>:
    name: 'pinec'
    canvas:
        Rectangle:
            source: resource_path('D:/Gugan/VOting system/bg.jpg')
            size: self.size 

    Label:
        text: 'Pine Captain'
        pos_hint: {'center_x':0.5,'center_y':0.88}
        font_size: 72
        bold: True

    Label:
        text:'NAME'
        pos_hint: {'center_x':0.3033,'center_y':0.38}
        color:0,0,1,1


    Label:
        text:'NAME'
        pos_hint: {'center_x':0.6075,'center_y':0.38}
        color:0,0,1,1

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.3033,'center_y':0.49}


    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.6075,'center_y':0.49}

    MDRaisedButton:
        id: but
        text: 'VOTE'
        pos_hint: {'center_x':0.3033,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1  
        on_press: app.pccontv1()
        disabled: False


    MDRaisedButton:
        id: but2
        text: 'VOTE'
        pos_hint: {'center_x':0.6075,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1
        on_press: app.pccontv2()
        disabled: False

<Pinev>:
    name: 'pinev'
    canvas:
        Rectangle:
            source: resource_path('D:/Gugan/VOting system/bg.jpg')
            size: self.size 

    Label:
        text: 'Pine Vice Captain'
        pos_hint: {'center_x':0.5,'center_y':0.88}
        font_size: 72
        bold: True

    Label:
        text:'NAME'
        pos_hint: {'center_x':0.3033,'center_y':0.38}
        color:0,0,1,1


    Label:
        text:'NAME'
        pos_hint: {'center_x':0.6075,'center_y':0.38}
        color:0,0,1,1

    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.3033,'center_y':0.49}


    Image:
        source: resource_path('D:/Gugan/VOting system/img.png')
        size_hint_x:0.18
        size_hint_y:0.17
        pos_hint: {'center_x':0.6075,'center_y':0.49}

    MDRaisedButton:
        id: but
        text: 'VOTE'
        pos_hint: {'center_x':0.3033,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1  
        on_press: app.pvcontv1()
        disabled: False


    MDRaisedButton:
        id: but2
        text: 'VOTE'
        pos_hint: {'center_x':0.6075,'center_y':0.331}
        size_hint_x:0.18
        opacity: 1
        on_press: app.pvcontv2()
        disabled: False
'''
kv1 = Builder.load_file('C:/Users/ADMIN/VOtingsystem/main.kv')

class MainApp (MDApp):
    wb = Workbook()
    ws = wb.active
    wb1 = Workbook()
    ws1 = wb1.active
    Maple = False
    Cedar = False
    Oak = False
    Pine = False
    no= 2
    def build(self):
        global sc
        sc = ScreenManager()
        sc.add_widget(LoginScreen(name = 'log'))
        sc.add_widget(PrefectScreen(name = 'votepb'))
        sc.add_widget(PrefectgScreen(name = 'votepg'))
        sc.add_widget(StudentScreen(name = 'votes'))
        sc.add_widget(Maple(name = 'maple'))
        sc.add_widget(Cedar(name = 'cedar'))
        sc.add_widget(Oak(name = 'oak'))
        sc.add_widget(Pinec(name = 'pinec'))
        sc.add_widget(Pinev(name = 'pinev'))
        return  sc
    def login(self,user,passw,error):
        adm = pandas.read_excel(resource_path('D:/Gugan/VOting system/Admin.xlsx'))
        vot = pandas.read_excel(resource_path('D:/Gugan/VOting system/voted.xlsx'))
        sadm = adm.values
        siadm = (sadm.size)
        usern = user.text
        passwo = passw.text
        
        for i in range(0,int((siadm)/3),1):
            
            y=  str(sadm[i][1]) 
            if passwo == y and usern in sadm[i] and int(passwo) not in vot.values:
                vbwb = Workbook()
                vbws = vbwb.active
                voload = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/novotes.xlsx'))
                voloada = voload.active
                voted = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/voted.xlsx'))
                voteda = voted.active
                #self.ws['A1'] = 'Name'
                #self.ws['B1'] = 'Password'
                #self.ws['A'+str(self.no)] = usern
                #self.ws['B'+str(self.no)] = passwo
                #self.wb.save(resource_path('D:/Gugan/VOting system/voted.xlsx'))
                for v in range(1,voloada.cell(row = 1 , column = 1).value+2):
                    self.ws['A'+str(v)] = voteda.cell(row = v, column = 1).value
                    self.ws['B'+str(v)] = voteda.cell(row = v, column = 2).value
                    self.wb.save(resource_path('D:/Gugan/VOting system/voted.xlsx'))
                self.ws['A'+str(voloada.cell(row = 1 , column = 1).value+2)] = usern
                self.ws['B'+str(voloada.cell(row = 1 , column = 1).value+2)] = passwo
                self.wb.save(resource_path('D:/Gugan/VOting system/voted.xlsx'))

                sc.current = 'votepb'
                error.text = ''
                if sadm[i][2] == 'Maple':
                    self.Maple = True
                if sadm[i][2] == 'Cedar':
                    self.Cedar = True
                if sadm[i][2] == 'Oak':
                    self.Oak = True
                if sadm[i][2] == 'Pine':
                    self.Pine = True
                self.no +=1
                vbws['A1'] = voloada.cell(row = 1, column = 1).value +1
                vbwb.save(resource_path('D:/Gugan/VOting system/novotes.xlsx'))
                break
            else:
                error.text = 'Wrong Credentials'
        

    def pbcontv1(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value + 1
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'votepg'
    def pbcontv2(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value + 1
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'votepg'
    def pbcontv3(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value + 1
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'votepg'

    def pgcontv1(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value + 1
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'votes'

    def pgcontv2(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value + 1
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'votes'

    def pgcontv3(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value + 1
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'votes'

    def pgcontv4(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value + 1
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'votes'


    def scontv1(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value + 1
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        if self.Maple == True:
            sc.current= 'maple'
            self.Maple = False
        if self.Cedar == True:
            sc.current= 'cedar'
            self.Cedar = False
        if self.Oak == True:
            sc.current= 'oak'
            self.Oak = False
        if self.Pine == True:
            sc.current= 'pinec'
            self.Pine = False

    def scontv2(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value + 1
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        if self.Maple == True:
            sc.current= 'maple'
            self.Maple = False
        if self.Cedar == True:
            sc.current= 'cedar'
            self.Cedar = False
        if self.Oak == True:
            sc.current= 'oak'
            self.Oak = False
        if self.Pine == True:
            sc.current= 'pinec'
            self.Pine = False      

    def scontv3(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value + 1
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        if self.Maple == True:
            sc.current= 'maple'
            self.Maple = False
        if self.Cedar == True:
            sc.current= 'cedar'
            self.Cedar = False
        if self.Oak == True:
            sc.current= 'oak'
            self.Oak = False
        if self.Pine == True:
            sc.current= 'pinec'
            self.Pine = False 

    def mcontv1(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value + 1
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'
        
    
    def mcontv2(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value + 1
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'
        
    
    def mcontv3(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value  + 1
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'
        
           

    def ccontv1(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value + 1
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value 
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'
        
    
    def ccontv2(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value + 1
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'
        
    
    def ccontv3(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value + 1
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'
        
           

    def ocontv1(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value + 1
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'
        
    
    def ocontv2(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value + 1
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'
        
    
    def ocontv3(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value + 1
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'
        
    

    def pccontv1(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value + 1
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'pinev'
    def pccontv2(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value + 1
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'pinev'

    def pvcontv1(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value + 1
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'
        
    
    def pvcontv2(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='pbcont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value 
        self.ws1['A2']='pbcont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='pbcont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='pgcont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='pgcont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='pgcont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='pgcont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='scont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value 
        self.ws1['A9']='scont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value 
        self.ws1['A10']='scont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value 
        self.ws1['A11']='mcont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value 
        self.ws1['A12']='mcont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value 
        self.ws1['A13']='mcont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value 
        self.ws1['A14']='ccont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value 
        self.ws1['A15']='ccont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='ccont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value 
        self.ws1['A17']='ocont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value 
        self.ws1['A18']='ocont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='ocont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pccont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pccont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pvcont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pvcont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value + 1
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'
        
    


if __name__ == "__main__":
    try:
        if hasattr(sys, "_MEIPASS"):
            resource_add_path(os.path.join(sys._MEIPASS))
        app = MainApp()
        app.run()
    except Exception as e:
        print(e)
        input("Press enter.")