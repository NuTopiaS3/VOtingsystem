import firebase_admin
from firebase_admin import credentials
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
storage = firebase.storage()

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
kv1 = Builder.load_file('D:/Gugan/VOting system/main.kv')

class MainApp(MDApp):
    current_admin_house = ''
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
        self.init_vote_functions()
        return self.sc

    def on_start(self):
        self.init_vote_functions()
        
    def init_vote_functions(self):
        categories = [
            ('cedar_captain_cont1', 'log'), ('cedar_captain_cont2', 'log'),
            ('cedar_captain_cont3', 'log'), ('cedar_captain_cont4', 'log'),
            ('maple_captain_cont1', 'log'), ('maple_captain_cont2', 'log'),
            ('maple_captain_cont3', 'log'), ('maple_captain_cont4', 'log'),
            ('oak_captain_cont1', 'log'), ('oak_captain_cont2', 'log'),
            ('oak_captain_cont3', 'log'), ('oak_captain_cont4', 'log'),
            ('pine_captain_cont1', 'pinev'), ('pine_captain_cont2', 'pinev'),
            ('pine_captain_cont3', 'pinev'), ('pine_captain_cont4', 'pinev'),
            ('pine_vc_cont1', 'log'), ('pine_vc_cont2', 'log'),
            ('pine_vc_cont3', 'log'), ('pine_vc_cont4', 'log'),
            ('prefect_boy_cont1', 'votepg'), ('prefect_boy_cont2', 'votepg'),
            ('prefect_boy_cont3', 'votepg'), ('prefect_boy_cont4', 'votepg'),
            ('prefect_girl_cont1', 'votes'), ('prefect_girl_cont2', 'votes'),
            ('prefect_girl_cont3', 'votes'), ('prefect_girl_cont4', 'votes'),
            ('student_council_cont1', 'votes'), ('student_council_cont2', 'votes'),
            ('student_council_cont3', 'votes'), ('student_council_cont4', 'votes'),
        ]
        for category, screen in categories:
            if category.startswith('student_council'):
                # For student council votes, call a special function that handles navigation based on the house
                setattr(self, category, lambda cat=category: self.handle_student_council_vote(cat))
            else:
                # For all other votes, use a direct navigation approach
                setattr(self, category, lambda cat=category, scr=screen, _self=self: _self.update_vote_count_and_navigate(cat, scr))

    def update_vote_count_and_navigate(self, category_key, next_screen):
        # Updates the vote count for the given category and navigates to the specified screen
        results = db.child("result").get().val()
        results[category_key] = results.get(category_key, 0) + 1
        db.child("result").set(results)
        print(f"Updating vote count for {category_key} and navigating to {next_screen}")
        self.sc.current = next_screen

    def handle_student_council_vote(self, category):
        # Special handling for student council voting to navigate based on house
        print(f"Handling student council vote for {category}")
        results = db.child("result").get().val()
        if category in results:
            results[category] += 1
        else:
            results[category] = 1
        db.child("result").update({category: results[category]})
        next_screen = self.get_next_screen_based_on_house()
        self.update_vote_count_and_navigate(category, next_screen)

    def get_next_screen_based_on_house(self):

        house_to_screen_mapping = {
            'Maple': 'maple',  # Assuming you have defined a screen named 'maple_screen'
            'Cedar': 'cedar',
            'Oak': 'oak',
            'Pine': 'pinec',
        }
        return house_to_screen_mapping.get(self.current_admin_house, 'default_screen')

    def get_current_admin_house(self):
        # Assuming you have a way to get the current user's ID
        current_user_id = firebase.auth().current_user.uid
        
        # Fetch the admin's data from the database
        admin_data = db.child("admins").child(current_user_id).get().val()
        
        # Extract and return the house information
        return admin_data.get('house', 'Unknown')

    def login(self, user, passw, error):
        # Simplified login logic with one-time voting check
        admins = db.child("Admin").get().val()
        voted = db.child("Voted").get().val()

        for index, admin in enumerate(admins[1:], start=1):  # Adjusted to match the indexing of your "Voted" list
            if admin["Name"] == user.text and admin["Password"] == passw.text:
                if not voted[index]:  # Check if the admin has not voted
                    self.current_admin_house = admin["House"]
                    # Mark the admin as having voted
                    db.child("Voted").child(str(index)).set(True)
                    # Proceed to the voting screen
                    self.sc.current = 'votepb'
                    print("Login Successful")
                    return
                else:
                    error.text = "You have already voted."
                    return
        error.text = "Login failed. Please check your credentials."

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
