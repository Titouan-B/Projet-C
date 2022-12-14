# -*- coding: utf-8 -*-
"""
Created on Wed Dec  7 15:13:05 2022

@author: boti4881
"""


import sys
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtCore import *
from PyQt6.QtCore import pyqtSignal as Signal, pyqtSlot as Slot
from packages import func

nb_doc=0


class Worker(QObject):
    progress = Signal(int)
    completed = Signal(int)

    @Slot(int)
    def do_work(self, n):
        global filedir
        ListeVin = pnnmoy['numvin'].tolist()
        for vin in enumerate(pnnmoy['numvin'].tolist()):
            winedata = pnntransformed.loc[pnntransformed["numvin"] == vin[1]]
            RPV = func.RadarplotVin(pnnmoy, vin[1])
            GraphPos = func.GraphPositionVin(df2, vin[1])
            GDC = func.GraphDemiCercle(winedata)
            text_output = []
            text_output = func.ConsoPourcentage(df3_count, vin[1], text_output)
            text_output = func.TranchePrix(df4transformed,vin[1],text_output)
            text_output += [str(chiffreannee)]
            func.CreaPDF(vin[1],text_output, filedir)
            self.progress.emit(int(vin[0]+1))
        else :
            self.completed.emit(int(vin[0]+1))



class MainWindow(QMainWindow):
    work_requested = Signal(int)

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        global chiffreannee
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
        layout.addWidget(self.label, 0,1)

        # Bouton pour chercher le fichier
        self.button1 = QPushButton("Importer fichier", clicked = self.Bouton_Import)
        layout.addWidget(self.button1, 0,0)


        # Radiobouton
        self.b1 = QRadioButton("Document conso", self)
        self.b1.setChecked(True)
        self.b1.setStyleSheet("QRadioButton { color : red; }")
        layout.addWidget(self.b1, 1,0)

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

    def update(self, value):
        global chiffreannee
        chiffreannee = value
    # multithread
    # Démarrage
    def start(self):
        self.button3.setEnabled(False)
        n = int(len(pnnmoy['numvin'].tolist()))
        self.prog_bar.setMaximum(n)
        self.work_requested.emit(n)

# Update de la barre
    def update_progress(self, v):
        self.prog_bar.setValue(v)
# Fin
    def complete(self, v):
        self.prog_bar.setValue(v)
        self.button3.setEnabled(True)


    # Bouton pour importer les excels
    def Bouton_Import(self):
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
            filename, _filter = QFileDialog.getOpenFileName(self, 'OpenFile')
            df, df2, pnnmoy, df3_count, df4transformed, pnntransformed = func.ImportData(filename)
            self.label.setText(filename[0:20] + '...')
            self.label.resize(200, 20)
            if self.b1.isChecked():
                self.b1.setStyleSheet("QRadioButton { color : pink; }")
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
            print('Mauvais input')
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
