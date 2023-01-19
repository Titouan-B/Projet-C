# -*- coding: utf-8 -*-
"""
Created on Wed Dec  7 15:13:05 2022

@author: boti4881
"""

# Import des modules
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtCore import *
from PyQt5.QtCore import pyqtSignal as Signal, pyqtSlot as Slot
from packages import func #func est un module à part contenant les fonctions de 
# stats et pour importer les données

# Compteur qui sert a suivre le nombre de documents chargés
nb_doc=0

# Worker = Multithreading = Programme qui fonctionne en parralèle d'un autre
# ici sert a créer les fiches 
class Worker(QObject):
    progress = Signal(int)
    completed = Signal(int)

    @Slot(int)
    # Fonction principale
    def do_work(self, n):
        global filedir
        # ListeVin = liste des numéros de vin
        ListeVin = pnnmoy['numvin'].tolist()
        # Pour chaque vin dans la liste, on crée les images/stats nécessaires 
        #  à mettre dans la fiche
        for vin in enumerate(pnnmoy['numvin'].tolist()):
            winedata = pnntransformed.loc[pnntransformed["numvin"] == vin[1]]
            # Radar plot
            RPV = func.RadarplotVin(pnnmoy, vin[1])
            # Graphique avec les positions des vins
            GraphPos = func.GraphPositionVin(df2, vin[1])
            # Graph en demi cercle
            GDC = func.GraphDemiCercle(winedata)
            # Tout le texte, donc pourcentage prix, tranches d'achats...
            text_output = []
            text_output = func.ConsoPourcentage(df3_count, vin[1], text_output)
            text_output = func.TranchePrix(df4transformed,vin[1],text_output)
            text_output += [str(chiffreannee)]
            # Fonction qui crée le pdf
            func.CreaPDF(vin[1],text_output, filedir)
            self.progress.emit(int(vin[0]+1))
        else :
            self.completed.emit(int(vin[0]+1))


# Classe qui sert à créer la fenetre principale
class MainWindow(QMainWindow):
    work_requested = Signal(int)

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Variable qui suit l'année 
        global chiffreannee
        # Titre de l'appli
        self.setWindowTitle("Application Eminents")
        self.widget = QWidget()
        # La fenetre est une "grille", càd que chaque élement a une position X
        #  et Y, en 0,0 en haut a gauche, X et Y étant des entiers
        layout = QGridLayout()

        # Mise en forme des colonnes et des rows
        for i in range(5):
            layout.setColumnMinimumWidth(max(i,3), 20)
            layout.setRowStretch(i, 1)

        # Taille de la fenetre
        self.resize(400, 200)
        self.widget.setLayout(layout)
        self.setCentralWidget(self.widget)

        # Label emplacement excel
        self.label = QLabel("...", self) #Ajoute du texte
        layout.addWidget(self.label, 0,1) #Ajoute ce texte en position 0,1

        # Bouton pour chercher le fichier
        self.button1 = QPushButton("Importer fichier", clicked = self.Bouton_Import)
        layout.addWidget(self.button1, 0,0)


        # Radiobouton pour selectionner quel document est importé
        self.b1 = QRadioButton("Document conso", self)
        self.b1.setChecked(True) #Coché par défaut
        self.b1.setStyleSheet("QRadioButton { color : red; }") #Rouge par défaut
        layout.addWidget(self.b1, 1,0) #Ajoute le bouton

        self.b2 = QRadioButton("Document pro", self)
        self.b2.setChecked(False)
        self.b2.setStyleSheet("QRadioButton { color : red; }")
        layout.addWidget(self.b2, 1,1)

        self.labelradio = QLabel("0/2 documents chargés", self)
        self.labelradio.setStyleSheet("QLabel { color : red; }")
        layout.addWidget(self.labelradio, 1,2)

        # Label dir pdf
        self.labeldir = QLabel("...", self)
        layout.addWidget(self.labeldir, 2,1)

        # Bouton dir pdf
        self.button2 = QPushButton("Emplacement du PDF", clicked = self.Bouton_Save)
        layout.addWidget(self.button2, 2,0)

        # Slider année éminents
        self.labelannee = QLabel("Année des éminents :", self)
        self.labelannee.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.labelannee, 3,0)
        annee = self.slider = QSpinBox( self, minimum=0, maximum=9999, value=2022)
        layout.addWidget(self.slider, 3,1)
        self.slider.setMinimum(2022)
        self.slider.setSingleStep(1)
        annee.valueChanged.connect(self.update)

        # gen PDF/Multithreading

        self.button3 = QPushButton("Créer les fiches PDF", clicked = self.start)
        layout.addWidget(self.button3, 4,0)


        # Loading bar
        self.prog_bar = QProgressBar(self)
        layout.addWidget(self.prog_bar, 5,0)
        self.prog_bar.setValue(0)

        # Tous les trucs pour lancer le worker et mettre les infos a jour
        self.worker = Worker()
        self.worker_thread = QThread()

        self.worker.progress.connect(self.update_progress)
        self.worker.completed.connect(self.complete)

        self.work_requested.connect(self.worker.do_work)

        # move worker to the worker thread
        self.worker.moveToThread(self.worker_thread)

        # start the thread
        self.worker_thread.start()

        # show the window
        self.show()

# Toutes les fonctions utilisées dans le programme : 

    
    def update(self, value):
        '''
        Met a jour le chiffre de l'année des éminents avec le slider'

        Parameters
        ----------
        value : str, année des éminents en sortie du slider

        Returns
        -------
        None.

        '''
        global chiffreannee
        chiffreannee = value
        
    # multithread
    # Démarrage
    def start(self):
        '''
        Démarre le mutlithreading du worker, désactive le bouton pour ça.
        Renvoie des valeurs pour update la barre de chargement

        Returns
        -------
        None.

        '''
        self.button3.setEnabled(False)
        n = int(len(pnnmoy['numvin'].tolist()))
        self.prog_bar.setMaximum(n)
        self.work_requested.emit(n)

# Update de la barre
    def update_progress(self, v):
        '''
        

        Parameters
        ----------
        v : int, nombre de pdfs créés

        Returns
        -------
        None.

        '''
        self.prog_bar.setValue(v)
# Fin
    def complete(self, v):
        '''
        Est utilisée quand le programme de création de pdf est terminée
        Réactive le bouton

        Parameters
        ----------
        v : int, nombre de pdfs créés

        Returns
        -------
        None.

        '''
        self.prog_bar.setValue(v)
        self.button3.setEnabled(True)


    # Bouton pour importer les excels
    def Bouton_Import(self):
        '''
        Importe les excel et montre combien sont actuellement importés

        Returns
        -------
        None.

        '''
        try :
            global annee
            global filename
            global nb_doc
            global pnnmoy
            global df3_count
            global df4transformed
            global df
            global df2
            global pnntransformed
            global chiffreannee
            chiffreannee = 2022
            
            # Ouvre une fenetre pour selectionner le fichier : 
            filename, _filter = QFileDialog.getOpenFileName(self, 'OpenFile')
            
            # Extrait toutes les données nécessaires : 
            df, df2, pnnmoy, df3_count, df4transformed, pnntransformed = func.ImportData(filename)
            
            self.label.setText(filename[0:20] + '...') #Affiche le lien du excel
            self.label.resize(200, 20)
            
            # Lorsqu'un doc est chargé, colore le bouton en vert et affiche le nb
            # de docs chargés à ce moments
            if self.b1.isChecked():
                self.b1.setStyleSheet("QRadioButton { color : green; }")
                nb_doc+=1
                self.labelradio.setText(str(nb_doc) + "/2 documents chargés")
                self.b2.setChecked(True)

            elif self.b2.isChecked():
                self.b2.setStyleSheet("QRadioButton { color : green; }")
                nb_doc+=1
                self.labelradio.setText(str(nb_doc) + "/2 documents chargés")
                self.b1.setChecked(True)

            if nb_doc==1:
                self.labelradio.setStyleSheet("QLabel { color : orange; }")
            if nb_doc==2:
                self.labelradio.setStyleSheet("QLabel { color : green; }")

        except:
            print('Mauvais input') #Si jamais le fichier importé n'est pas valide
            pass
        
    # Bouton pour importer demander ou sauvegarder le pdf
    def Bouton_Save(self):
        global filedir
        filedir = QFileDialog.getExistingDirectory(self)
        self.labeldir.setText(filedir[0:20] + '...')
        self.labeldir.resize(200, 20)
        return filedir

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    window = MainWindow()
    sys.exit(app.exec())
