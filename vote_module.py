from kivymd.app import MDApp
from kivy.lang import Builder
from kivy.uix.screenmanager import Screen,ScreenManager
from kivy.animation import Animation
import pandas
from openpyxl import Workbook
import openpyxl
import os,sys
from kivy.resources import resource_add_path, resource_find
from kivy.core.window import Window
from kivy.uix.floatlayout import FloatLayout
import pyrebase

# Firebase configuration
config = {
   "apiKey": "apiKey",
   "authDomain": "projectId.firebaseapp.com",
   "databaseURL": "https://voting-system-gugan-default-rtdb.asia-southeast1.firebasedatabase.app/",
   "projectId": "projectId",
   "storageBucket": "projectId.appspot.com",
   "messagingSenderId": "messagingSenderId",
   "appId": "appId"
}

# Instantiates a Firebase app
app = pyrebase.initialize_app(config)


# Firebase Realtime Database
db = app.database()

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


kv='''s
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
#this alone for else then it works perfect delete the below code later                error.text = 'Wrong Credentials'
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
#delete thisss


    def prefect_boy_cont1(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value + 1
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'votepg'
    def prefect_boy_cont2(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value + 1
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'votepg'
    def prefect_boy_cont3(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value + 1
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'votepg'

    def prefect_girl_cont1(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value + 1
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'votes'

    def prefect_girl_cont2(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value + 1
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'votes'

    def prefect_girl_cont3(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value + 1
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'votes'

    def prefect_girl_cont4(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value + 1
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'votes'


    def student_council_cont1(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value + 1
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
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

    def student_council_cont2(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value + 1
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
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

    def student_council_cont3(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value + 1
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
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

    def maple_captain_cont1(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value + 1
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'


    def maple_captain_cont2(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value + 1
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'


    def maple_captain_cont3(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value  + 1
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'



    def cedar_captain_cont1(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value + 1
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'


    def cedar_captain_cont2(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value + 1
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'


    def cedar_captain_cont3(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value + 1
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'



    def oak_captain_cont1(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value + 1
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'


    def oak_captain_cont2(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value + 1
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'


    def oak_captain_cont3(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value + 1
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'



    def pine_captain_cont1(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value + 1
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'pinev'
    def pine_captain_cont2(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value + 1
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'pinev'

    def pine_vc_cont1(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value + 1
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'


    def pine_vc_cont2(self):
        load = openpyxl.load_workbook(resource_path('D:/Gugan/VOting system/result.xlsx'))
        loada = load.active
        self.ws1['A1']='prefect_boy_cont1: '
        self.ws1['B1']= loada.cell(row = 1, column = 2).value
        self.ws1['A2']='prefect_boy_cont2: '
        self.ws1['B2']=loada.cell(row = 2, column = 2).value
        self.ws1['A3']='prefect_boy_cont3: '
        self.ws1['B3']=loada.cell(row = 3, column = 2).value
        self.ws1['A4']='prefect_girl_cont1: '
        self.ws1['B4']=loada.cell(row = 4, column = 2).value
        self.ws1['A5']='prefect_girl_cont2: '
        self.ws1['B5']=loada.cell(row = 5, column = 2).value
        self.ws1['A6']='prefect_girl_cont3: '
        self.ws1['B6']=loada.cell(row = 6, column = 2).value
        self.ws1['A7']='prefect_girl_cont4: '
        self.ws1['B7']=loada.cell(row = 7, column = 2).value
        self.ws1['A8']='student_council_cont1 : '
        self.ws1['B8']=loada.cell(row = 8, column = 2).value
        self.ws1['A9']='student_council_cont2 : '
        self.ws1['B9']=loada.cell(row = 9, column = 2).value
        self.ws1['A10']='student_council_cont3 : '
        self.ws1['B10']=loada.cell(row = 10, column = 2).value
        self.ws1['A11']='maple_captain_cont1 : '
        self.ws1['B11']=loada.cell(row = 11, column = 2).value
        self.ws1['A12']='maple_captain_cont2 : '
        self.ws1['B12']=loada.cell(row = 12, column = 2).value
        self.ws1['A13']='maple_captain_cont3 : '
        self.ws1['B13']=loada.cell(row = 13, column = 2).value
        self.ws1['A14']='cedar_captain_cont1 : '
        self.ws1['B14']=loada.cell(row = 14, column = 2).value
        self.ws1['A15']='cedar_captain_cont2 : '
        self.ws1['B15']=loada.cell(row = 15, column = 2).value
        self.ws1['A16']='cedar_captain_cont3 : '
        self.ws1['B16']=loada.cell(row = 16, column = 2).value
        self.ws1['A17']='oak_captain_cont1 : '
        self.ws1['B17']=loada.cell(row = 17, column = 2).value
        self.ws1['A18']='oak_captain_cont2 : '
        self.ws1['B18']=loada.cell(row = 18, column = 2).value
        self.ws1['A19']='oak_captain_cont3 : '
        self.ws1['B19']=loada.cell(row = 19, column = 2).value
        self.ws1['A20']='pine_captain_cont1: '
        self.ws1['B20']=loada.cell(row = 20, column = 2).value
        self.ws1['A21']='pine_captain_cont2: '
        self.ws1['B21']=loada.cell(row = 21, column = 2).value
        self.ws1['A22']='pine_vc_cont1: '
        self.ws1['B22']=loada.cell(row = 22, column = 2).value
        self.ws1['A23']='pine_vc_cont2: '
        self.ws1['B23']=loada.cell(row = 23, column = 2).value + 1
        self.wb1.save(resource_path('D:/Gugan/VOting system/result.xlsx'))
        sc.current = 'log'


Window.fullscreen = True
Window.size=(1920,1080)

if __name__ == "__main__":
    try:
        if hasattr(sys, "_MEIPASS"):
            resource_add_path(os.path.join(sys._MEIPASS))
        app = MainApp()
        app.run()
    except Exception as e:
        print(e)
        input("Press enter.")
