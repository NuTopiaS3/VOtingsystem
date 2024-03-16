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
    "apiKey": "AIzaSyDy-x6jIFpY-Kb4rsCk1-0SRTwSQlD2R3E",
    "authDomain": "voting-system-gugan.firebaseapp.com",
    "databaseURL": "https://voting-system-gugan-default-rtdb.asia-southeast1.firebasedatabase.app",
    "projectId": "voting-system-gugan",
    "storageBucket": "voting-system-gugan.appspot.com",
    "messagingSenderId": "1044881782390",
    "appId": "1:1044881782390:web:571667c95213fc7a5f9157"
}
# Initialize Firebase app
firebase = pyrebase.initialize_app(config)
db = firebase.database()

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

class MainApp(MDApp):
    def build(self):
        self.sc = ScreenManager()
        self.sc.add_widget(LoginScreen(name='log'))
        self.sc.add_widget(PrefectScreen(name='votepb'))
        self.sc.add_widget(PrefectgScreen(name='votepg'))
        self.sc.add_widget(StudentScreen(name='votes'))
        self.sc.add_widget(Maple(name='maple'))
        self.sc.add_widget(Cedar(name='cedar'))
        self.sc.add_widget(Oak(name='oak'))
        self.sc.add_widget(Pinec(name='pinec'))
        self.sc.add_widget(Pinev(name='pinev'))
        return self.sc
    def login(self, user, passw, error):
        username = user.text
        password = passw.text

        # Retrieve admin data from Firebase
        admin_data_list = db.child("Admin").get().val()

        print("Admin data retrieved:", admin_data_list)  # Debugging

        if admin_data_list:
            print("Type of admin_data:", type(admin_data_list))

            # Filter out None elements from the list
            admin_data_list = [admin_info for admin_info in admin_data_list if admin_info is not None]

            # Iterate over the list of admins
            for admin_info in admin_data_list:
                if admin_info["Name"] == username and admin_info["Password"] == password:
                    admin_id = str(admin_info["ID"])

                    # Check if the user has already voted
                    voted_status = db.child("Voted").child(admin_id).get().val()
                    if voted_status:
                        error.text = 'You have already voted'
                    else:
                        house = admin_info["House"]
                        # Update voting status in Firebase
                        db.child("Voted").child(admin_id).set(True)

                        # Navigate to appropriate screen based on house
                        if house == 'Maple':
                            self.sc.current = 'maple'
                        elif house == 'Cedar':
                            self.sc.current = 'cedar'
                        elif house == 'Oak':
                            self.sc.current = 'oak'
                        elif house == 'Pine':
                            self.sc.current = 'pinev'

                        error.text = ''
                        break
            else:
                error.text = 'Wrong Credentials'
        else:
            error.text = 'No admin data found or error fetching data from Firebase'
    def prefect_boy_cont1(self):
        results = db.child("result").get().val()
        results['prefect_boy_cont1'] = results.get('prefect_boy_cont1', 0) + 1
        db.child("result").set(results)
        self.sc.current = 'votepg'
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
