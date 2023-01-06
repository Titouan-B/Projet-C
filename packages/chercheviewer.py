# -*- coding: utf-8 -*-
"""
Created on Thu Dec 22 18:15:14 2022

@author: boti4881
"""

import requests, json

import sys
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtCore import *



class MainWindow(QMainWindow):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setWindowTitle("Application Eminents")
        self.widget = QWidget()
        layout = QGridLayout()

        for i in range(5):
            layout.setColumnMinimumWidth(max(i,3), 20)
            layout.setRowStretch(i, 1)

        self.resize(400, 200)
        self.widget.setLayout(layout)
        self.setCentralWidget(self.widget)

        # Label emplacement excel
        self.label = QLabel("...", self)
        layout.addWidget(self.label, 2,1)

        # Bouton pour chercher le fichier
        self.button1 = QPushButton("Chercher viewer", clicked = self.Cherche)
        layout.addWidget(self.button1, 2,0)

        # Nom viewer
        self.labelv = QLabel("Nom viewer : ", self)
        layout.addWidget(self.labelv, 1,0)
        
        self.line = QLineEdit(self)
        layout.addWidget(self.line, 1,1)
        
        
        self.labels = QLabel("Nom streamer : ", self)
        layout.addWidget(self.labels, 0,0)
            
        self.button2 = QPushButton("Rafraichir", clicked = self.Refresh)
        layout.addWidget(self.button2, 0,2)
        self.labeln = QLabel("Nb viewers : ?", self)
        layout.addWidget(self.labeln, 1,2)
        # Nom streamer
        self.line2 = QLineEdit(self)
        layout.addWidget(self.line2, 0,1)
        
        # show the window
        self.show()

    def Cherche(self):
        response = requests.get("https://tmi.twitch.tv/group/user/" + self.line2.text().lower().replace(" ", "") + "/chatters").text
        response_info = json.loads(response)
        self.labeln.setText("Nb viewers : "+str(response_info['chatter_count']))
        for key in response_info['chatters'].keys():
            if self.line.text() in response_info['chatters'][key]:
                self.label.setText('Il/Elle regarde !')
                break
            else :
                self.label.setText('Il/Elle ne regarde pas...')
            
    def Refresh(self):
        response = requests.get("https://tmi.twitch.tv/group/user/" + self.line2.text().lower().replace(" ", "") + "/chatters").text
        response_info = json.loads(response)
        self.labeln.setText("Nb viewers : "+str(response_info['chatter_count']))
        
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    sys.exit(app.exec())



# while True:
#     response = requests.get("https://tmi.twitch.tv/group/user/ponce/chatters").text
#     response_info = json.loads(response)
#     if 'tigreed' in response_info['chatters']['viewers']:
#         print('Titou regarde !')
