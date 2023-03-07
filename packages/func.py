# -*- coding: utf-8 -*-
"""
Created on Fri Nov 25 22:12:03 2022

@author: boti4881
"""



# Slider année


import matplotlib.pyplot as plt
import pandas as pd
from seaborn import kdeplot, set
import numpy as np
import fitz
from os import chdir, getcwd

try:
    import pyi_splash

    pyi_splash.close()
except ModuleNotFoundError:
    pass

#Script test pour demander un vin : changer le numéro pour changer de vin
# vin = 109


# =============================================================================

# Traitement/Mise en forme

def ImportDataConso(filename):
    '''
    Fonction important et traitant les données

    Parameters
    ----------
    filename : str
        Filename of the excel file

    Returns
    -------
    None.

    '''
#Mise en forme des dataframe a partir des fichiers
    df = pd.read_excel(filename)
    df2 = df.groupby(["numvin"]).mean()

# df3 -> colonnes avec conditions de consommation
    df3 = df.loc[:, 'q13_1':'q13_5'].copy()
    df3transformed = df['numvin'].copy()

# df4 -> colonnes avec tranches de prix
    df4 = df['q14'].copy()
    df4transformed = df['numvin'].copy()

# On sépare les colonnes avec positivement/négativement/pas d'influence
    pnn = df.loc[:, 'q3':'q11'].copy()

    pnntransformed = df['numvin'].copy()

#On fait une loop pour map une a une les colonnes
    for i in range(len(pnn.keys())):
        loop = pnn[[pnn.keys()[i]]].squeeze()
        loop = loop.map({'Positivement': int(1), "Pas d'influence": int(0), 'Négativement':int(-1) })
        pnntransformed = pd.concat([pnntransformed,loop], axis= 1)

    for i in range(len(df3.keys())):
        loop = df3[[df3.keys()[i]]].squeeze()
        loop = loop.map({'A boire en cours de repas': int(1), "Pour un simple moment festif": int(1), 'Parfait pour l’apéritif':int(1), 'Parfait pour le dessert':int(1),'Pour célébrer les grandes étapes de la vie':int(1) })
        df3transformed = pd.concat([df3transformed,loop], axis= 1)

    loop = df4.squeeze()
    loop = loop.map({'10 à 15 €': int(1), '15 à 20 €': int(0), '20 à 25 €':int(-1), '25 € ou plus':int(2) })
    df4transformed = pd.concat([df4transformed,loop], axis= 1)
# On obtient pnn transformed : la dataframe comportant des
# valeurs numériques (1/0/-1), on peut y appliquer des operations avec groupby

# pnn moy : la moyenne de chaque vin sur ces colonnes
    pnnmoy = pnntransformed.groupby("numvin").mean().reset_index()



# df3 count : comptage des réponses pour chaque vin et chaque type de conso
    df3_count = df3transformed.groupby("numvin").count()
    df3_count = df3_count.reset_index()


# df4 count : comptage des réponses pour chaque vin et chaque prix
    # df4data = df4transformed.loc[df4transformed["numvin"] == vin]

# toutes les données pour chaque question pour un vin
    # winedata = pnntransformed.loc[pnntransformed["numvin"] == vin]

    return df, df2, pnnmoy, df3_count, df4transformed, pnntransformed
# Ajouter return df4 et tout, verifier que tout colle dans start au niveau
# des retours, ajouter conso %

# =======================================================================

def ImportDataPro(filename):
    '''
    Fonction important et traitant les données
    Parameters
    ----------
    filename : str
        Filename of the excel file
    Returns
    -------
    None.
    '''

    dp =pd.read_excel(filename)
    dp2 = dp.groupby(["numvin"]).mean()

    dp3 = dp.loc[:, 'q15_1':'q15_5'].copy()
    dp3transformed = dp['numvin'].copy()

    dp4 = dp['q16'].copy()
    dp4transformed = dp['numvin'].copy()

    # Pour les pros on fait la moyenne des notes hédonniques pour chaque vin

    notes_hedoniques=pd.DataFrame(dp2.iloc[:, 1:14])
    moy_hedo_pro=pd.DataFrame(notes_hedoniques.groupby(["numvin"]).mean())


    for i in range(len(dp3.keys())):
        loop = dp3[[dp3.keys()[i]]].squeeze()
        loop = loop.map({'A boire en cours de repas': int(1), "Pour un simple moment festif": int(1), 'Parfait pour l’apéritif':int(1), 'Parfait pour le dessert':int(1),'Pour célébrer les grandes étapes de la vie':int(1) })
        dp3transformed = pd.concat([dp3transformed,loop], axis= 1)
    
    
    loop = dp4.squeeze()
    loop = loop.map({'10 à 15€': int(1), '15 à 20€': int(0), '20 à 25€':int(-1), '25€ ou plus':int(2) })
    dp4transformed = pd.concat([dp4transformed,loop], axis= 1)

    dp3_count = dp3transformed.groupby("numvin").count()
    dp3_count = dp3_count.reset_index()

    return dp, dp2, dp3_count, dp4transformed, moy_hedo_pro


# =============================================================================


def ConsoPourcentage(df3_count, vin, text_output):
    
    countvin = df3_count.loc[df3_count["numvin"]== vin]
    total = int(countvin['q13_1'] + countvin['q13_2']+ countvin['q13_3'] + countvin['q13_4'] + countvin['q13_5'])
    
    for col in enumerate(countvin.keys()[1:6].tolist()):
        text_output.append(str(int((countvin[col[1]]/total)*100)) + "%")
    return(text_output)

# =============================================================================


def TranchePrix(df4transformed,vin,text_output):
    df4data = df4transformed.loc[df4transformed["numvin"] == vin]

    dixquinze = df4data['q14'].loc[df4data['q14'] == 1].count()
    quinzevingt = df4data['q14'].loc[df4data['q14'] == 0].count()
    vingtvingcinq = df4data['q14'].loc[df4data['q14'] == -1].count()
    vingcinqplus = df4data['q14'].loc[df4data['q14'] == 2].count()
    
    total = dixquinze + dixquinze + vingtvingcinq + vingcinqplus

    text_output.append(str(int(dixquinze/total*100)))
    text_output.append(str(int(quinzevingt/total*100)))
    text_output.append(str(int(vingtvingcinq/total*100)))
    text_output.append(str(int(vingcinqplus/total*100)))
    return(text_output)

# =============================================================================


# Nuage de mots 

# from wordcloud import WordCloud
# import numpy as np
# from spellchecker import SpellChecker
    

# # Faire une longue liste de tous les mots

# words = df['q12'].loc[df["numvin"] == 209]
# words = words.dropna()
# wordlist = ""
# for i in range(len(words)):
#     wordlist += words.iloc[i]
#     wordlist += " "
    
# wordlist = wordlist.lower()


# ============ Vérificateur orthographique ? Marche pas vrmt j'ai l'impression
# spell = SpellChecker(language='fr')

# # find those words that may be misspelled
# misspelled = spell.unknown(wordlist)

# for word in misspelled:
#     # Get the one `most likely` answer
#     print(spell.correction(word))
# =============


# stopwords = ['vin','pa','l','d','moi','ni','sur','ai','plus','moins','plutôt','les','donne','ne','pourtant','avec','j','pour','vu','mon','trouve','malgré','ce','par','à','de','des','il','au', 'le', 'que', 'mais', 'la', 'un', 'et','trop','pas', 'du', 'je', 'peu','très','est','car','assez','en']

# cloud_generator = WordCloud(width=800, height=400, background_color=None,
#                             random_state=1, max_words=20,prefer_horizontal=1,
#                             stopwords = stopwords, mode='RGBA', colormap = 'PiYG',
#                             )

# wordcloud_image = cloud_generator.generate(wordlist)
# plt.figure( figsize=(20,10), facecolor='k' )
# plt.imshow(wordcloud_image, interpolation="bilinear")
# plt.savefig('w.png', format='png')
# plt.show()
# =============================================================================

def GraphDemiCercle(winedata):
    plt.clf() #Clear previous graphs in memory

# Labels supprimés a l'avenir pour une légende globale ?
    # label = ["Positif", "Neutre", "Négatif"]
    label = ["","",""]
#Compter le nombre de valeurs positives/neutres/négatives
    pos = winedata['q3'].loc[winedata['q3'] == 1].count()
    neu = winedata['q3'].loc[winedata['q3'] == 0].count()
    neg = winedata['q3'].loc[winedata['q3'] == -1].count()
    
    val = [pos,neu,neg]

# ajouter les données et la couleur
    label.append("")
    val.append(sum(val))
    colors = ["#cea450", '#727272','#030303', 'k']

# plot + titre
    fig = plt.figure(figsize=(8,6),dpi=100)
    wedges, labels = plt.pie(val, wedgeprops=dict(width=0.4,edgecolor='k'),labels=label, colors=colors)
    wedges[-1].set_visible(False)
    # plt.title('Comment votre bouteille a affecté les consos', fontfamily = "Arial Rounded MT Bold")
    plt.savefig('f.png', format='png', transparent=True, dpi=100)
    plt.close()
    return()

# =============================================================================



def GraphPositionVinNote(df2, moy_hedo_pro, vin):
# Graph du vin par rapport aux autres
    plt.clf() #Clear previous graphs in memory
    custom_params = {"axes.spines.right": False, "axes.spines.top": False}
    set(style="ticks", rc=custom_params)
    # Mise en forme des données

    rdfconso = df2['q2'].reset_index()
    rdfpro = moy_hedo_pro['q1'].reset_index()
    
    # Faire le plot + ligne du vin
    meanq1 = rdfpro['q1'].loc[rdfpro['numvin'] == vin]
    meanq2 = rdfconso['q2'].loc[rdfconso['numvin'] == vin]
    
    # Lignes insiquant la position du vin
    plt.xlim(xmax = 10, xmin = 0)
    plt.axvline(float(meanq1), 0,0.95, color = 'gold',ls = '--')
    # plt.text(float(meanq2)-0.4, 0.6,"Votre vin", fontfamily = "Arial Rounded MT Bold")
    plt.axvline(float(meanq2), 0,0.95, color = 'black', ls = '--')
    plt.xlabel("Note moyenne du crémant")
    plt.ylabel("Quantité de crémants (%)")
    # Faire les deux plots + titre et enlever l'axe y
    fig2 = kdeplot(rdfpro['q1'], shade=True, bw_method=0.5, color="gold")
    fig2 = kdeplot(rdfconso['q2'], shade=True, bw_method=0.5, color="black")
    fig2.figure.suptitle('Votre vin par rapport aux autres !', fontfamily = "Arial Rounded MT Bold")
    fig2.get_yaxis().set_visible(False)
    plt.savefig('pn.png', format='png', transparent=True, dpi=100)
    plt.close()
    return()

# =============================================================================



def GraphPositionVinScore(df2, moy_hedo_pro, vin):
# Graph du vin par rapport aux autres
    # Graph du vin par rapport aux autres
    plt.clf() #Clear previous graphs in memory
    custom_params = {"axes.spines.right": False, "axes.spines.top": False}
    set(style="ticks", rc=custom_params)
        # Mise en forme des données

    rdfconso = df2['q2'].reset_index()
    rdfpro = moy_hedo_pro['q1'].reset_index()
        
        # Faire le plot + ligne du vin
    meanq1 = rdfpro['q1'].loc[rdfpro['numvin'] == vin]
    # meanq2 = rdfconso['q2'].loc[rdfconso['numvin'] == vin]
        
        # Lignes insiquant la position du vin
    plt.xlim(xmax = 10, xmin = 0)
    plt.axvline(float(meanq1), 0,0.95, color = 'gold',ls = '--')
    # plt.text(float((meanq2+meanq1)/2)-0.6, 0.5,"Votre vin", fontfamily = "Arial Rounded MT Bold")
    # plt.axvline(float(meanq2), 0,0.95, color = 'black', ls = '--')
    plt.xlabel("Score moyen du crémant")
    plt.ylabel("Quantité de crémants (%)")
        # Faire les deux plots + titre et enlever l'axe y
    fig2 = kdeplot(rdfpro['q1'], shade=True, bw_method=0.5, color="gold")
    # fig2 = kdeplot(rdfconso['q2'], shade=True, bw_method=0.5, color="black")
    fig2.figure.suptitle('Votre vin par rapport aux autres !', fontfamily = "Arial Rounded MT Bold")
    fig2.get_yaxis().set_visible(True)
    plt.yticks(fig2.get_yticks(), np.around(fig2.get_yticks() * 100,0))

    # plt.yticks([])
    plt.savefig('ps.png', format='png', transparent=False, dpi=200)
    plt.close()
    return()

# =============================================================================


# #  Histogramme des notes - A voir avec Dujourdy & co

# from math import pi

# fig, ax = plt.subplots()
# sns.histplot(df['q2'].loc[df['numvin'] == vin],ax=ax, color = 'gold')
# ax.set_xlim(0,11)
# ax.set_xticks(range(0,11))
# ax.get_yaxis().set_visible(False)
# plt.show()


# =============================================================================



##Radarplot test

def RadarplotVin(dp, dp2, vin, j):

    
    dp2 = dp2.reset_index()
    maxval = []
    minval = []
    categories = []
    data = []
    moy = []
    # Echelle du radarplot
    markers = [0, 1, 2, 3, 4, 5, 6]

    # Catégories présentes sur le graphique (exemple ici)
    if j==0:
        # categories = ["Mousse", "Train de bulles", "Finesse des bulles", "Robe","Tenue cordon"]
        # categories = [*categories, categories[0]]
        
        data = dp2.loc[dp2["numvin"] == vin].values.tolist()[0][3:8]
        moy = dp2.mean()[3:8]
        data = [*data, data[0]]
        moy = [*moy, moy[0]]

        
       
        questions = dp.keys()[3:8].values.tolist() #on prend les noms des colonnes
        for i in range(len(questions)):
            maxval.append(max(dp2[questions[i]]))
            minval.append(min(dp2[questions[i]]))
        
        # Points max et min

        minval = [*minval, minval[0]]
        maxval = [*maxval, maxval[0]]
        
    if j==1:
        # categories = ["Intensité", "Netteté", "Nuances odorantes"]
        # categories = [*categories, categories[0]]
        
        data = dp2.loc[dp2["numvin"] == vin].values.tolist()[0][8:11]
        moy = dp2.mean()[8:11]
        data = [*data, data[0]]
        moy = [*moy, moy[0]]

        # Points max et min

         
        questions = dp.keys()[8:11].values.tolist() #on prend les noms des colonnes
        for i in range(len(questions)):
            maxval.append(max(dp2[questions[i]]))
            minval.append(min(dp2[questions[i]]))
        
        minval = [*minval, minval[0]]
        maxval = [*maxval, maxval[0]]
        
    if j==2:
        # categories = ["Effervescence", "Equilibre", "Finale"]
        # c = [*categories, categories[0]]
    
        data = dp2.loc[dp2["numvin"] == vin].values.tolist()[0][11:14]
        moy = dp2.mean()[11:14]
        data = [*data, data[0]]
        moy = [*moy, moy[0]]
    
        # Points max et min

         
        questions = dp.keys()[11:14].values.tolist() #on prend les noms des colonnes
        for i in range(len(questions)):
            maxval.append(max(dp2[questions[i]]))
            minval.append(min(dp2[questions[i]]))

        minval = [*minval, minval[0]]
        maxval = [*maxval, maxval[0]]

    

    # Gestion des angles et plot
    label_loc = np.linspace(start=0, stop=2 * np.pi, num=len(data))
    plt.figure(figsize=(8, 8))
    ax = plt.subplot(polar=True) #le polar fait que c'est un graph radial

    # Supprimer lex axes pour mettre devant un template sur le pdf ?
    # ax.axis('on')

    # Points moy
    plt.plot(label_loc, moy,'--', label=str(vin), color='#C0C0C0',linewidth=4.0, ms=30)
    plt.scatter(label_loc, moy, color='#C0C0C0',s=150)

    # Lignes + points du vin actuel
    plt.plot(label_loc, data,'o-', label=str(vin), color='#cea450',linewidth=4.0, ms=30)

    plt.scatter(label_loc, data, color='#dfbe62',s=200,zorder=-1)


    # Points max/min
    plt.scatter(label_loc, maxval, color='#030303',s=300, zorder=10)
    plt.scatter(label_loc, minval, color='#727272',s=300, zorder=10)

    #Plot + titre
    plt.yticks(markers)
    lines, labels = plt.thetagrids(np.degrees(label_loc))#, labels=questions)
    # axes.spines['polar'].set_visible(False)
    ax.set_xticklabels([])
    plt.savefig('r' + str(j) + '.png', format='png', transparent=True, dpi=400)
    plt.close()


    return()
    

# =============================================================================

def CalculScore(df2 ,dp2, pnnmoy, vin, i):
    # Calcul score :
    dp2 = dp2.reset_index()
    dp2temp = dp2.loc[dp2["numvin"] == vin]
    #SENSORIEL PRO 
    oeil = 0.15*dp2temp['q2'] + 0.325 * dp2temp['q5'] + 0.15*dp2temp['q3'] + 0.325 *dp2temp['q4'] + 0.15*dp2temp['q6']
    bouche = 0.2*dp2temp['q10'] + 0.4*dp2temp['q11'] + 0.4*dp2temp['q12']
    nez = 0.2*dp2temp['q7'] + 0.4*dp2temp['q9'] + 0.4*dp2temp['q8']
    ensemble = dp2temp['q13']
    SENSO = (0.1 * oeil + 0.3* bouche +nez *0.3 + ensemble *0.3)/6

    # ESTHETIQUE CONSO
    pnntemp = pnnmoy.loc[pnnmoy["numvin"] == vin]

    contenant = 0.6*pnntemp['q3'] + 0.4 *pnntemp['q4']
    coiffecol = pnntemp['q5']
    ensemblebout = pnntemp['q11']
    etiqcontreetiq = 0.3*pnntemp['q7'] + 0.15*pnntemp['q6'] + 0.2*pnntemp['q8'] + 0.2*pnntemp['q9'] +0.15*pnntemp['q10']

    ESTHETIQUE = (0.2 * contenant + 0.1*coiffecol + 0.4 * ensemblebout + 0.3 * etiqcontreetiq+1)/2

    # NOTE HEDO
    df2temp = df2.reset_index()
    df2temp = df2temp.loc[df2temp["numvin"] == vin]

    HEDO = (1/3 * dp2temp['q1'] + 1/3 * df2temp['q1'] + 1/3 * df2temp['q2'])/10

    # NOTE FINALE

    SCORE = 1/3 * SENSO + 1/3 * ESTHETIQUE + 1/3 * HEDO
    

    # On crée le graph
    pie = [SENSO[i],ESTHETIQUE[i],HEDO[i]]
    color = ['#727272', '#545454', '#343434']
    # plt.pie(pie, labels = ['Sensoriel (évaluation pro)', 'Esthétique (évaluation conso)', 'Note hédonique'],colors = color, normalize = True, autopct = lambda pie: str(round(pie, 1)) + '%',textprops={'color':"w"}, pctdistance=0.75)

    _, _, autotexts = plt.pie(pie, labels = ['Sensoriel \n(évaluation pro)', 'Esthétique \n(évaluation conso)', 'Note hédonique'],colors = color, normalize = True, autopct = lambda pie: str(round(pie, 1)) + '%', pctdistance=0.78)
    for autotext in autotexts:
        autotext.set_color('white')

    hole = plt.Circle((0, 0), 0.50, facecolor='white')
    plt.gcf().gca().add_artist(hole)
    
    plt.savefig('s.png', format='png', transparent=True, dpi=400)
    plt.close()
    return SCORE

# =============================================================================
# Modifier pdf dans Python
import os

def CreaPDF(vin,filedir, currdir,text_output,df2,dp2,i, doc):
    chdir(filedir)
    # page = doc.new_page()
    
    df2temp = df2.reset_index()
    df2temp = df2temp.loc[df2temp["numvin"] == vin]
    
    dp2temp = dp2.reset_index()
    dp2temp = dp2temp.loc[dp2temp["numvin"] == vin]
    
    for page in doc:
        page.clean_contents()  
        
    # PAGE 1
    page = doc.load_page(0)
    
    page.insert_textbox(fitz.Rect(0, 10, 595, 20), buffer = text_output[-1], fontname ="Rosmatika", fontfile = currdir + "Fonts//Rosmatika (DEMO).ttf",  align = 1, fontsize = 20)
    
    page.insert_textbox(fitz.Rect(0, 70, 595, 120), "Louis Bouillot", fontname ="/Rosmatika",   align = 1, fontsize = 50, color = (0.75, 0.62, 0.43))
    page.insert_textbox(fitz.Rect(0, 120, 595, 170), "Grand terroir", fontname ="/Rosmatika",    align = 1, fontsize = 50)
    page.insert_textbox(fitz.Rect(0, 170, 595, 210), "Chenôvre Hermitage", fontname ="/Rosmatika",    align = 1, fontsize = 50)
    page.insert_textbox(fitz.Rect(0, 210, 595, 250), buffer = "(" + str(vin) + ")", fontname ="Anton",  fontfile = currdir + "Fonts//Anton-Regular.ttf",   align = 1, fontsize = 20)
        
    page.insert_textbox(fitz.Rect(148, 326, 595, 380), buffer = str(round(dp2temp['q1'][i],2)),fontname ="/Anton", align = 1, fontsize = 21.4)

    
    page.insert_image(fitz.Rect(0,365,190,550), filename='r0.png')
    page.insert_image(fitz.Rect(198,365,386,550), filename='r1.png')
    page.insert_image(fitz.Rect(404,365,594,555), filename='r2.png')

    fontsize = 12+ np.log(max(int(text_output[5]),1)**2)
    page.insert_textbox(fitz.Rect(48, 790 + 1/fontsize * 60, 100, 830), buffer = text_output[5]  + '€', fontname ="/Rosmatika",   align = 1, fontsize = fontsize)

    fontsize = 12+ np.log(max(int(text_output[6]),1)**2)
    page.insert_textbox(fitz.Rect(45 + 140, 790+ 1/fontsize * 60 , 100 + 137, 830), buffer = text_output[6] + '€', fontname ="/Rosmatika",   align = 1, fontsize = 12+ np.log(max(int(text_output[6]),1)**2))
        
    fontsize = 12+ np.log(max(int(text_output[7]),1)**2)
    page.insert_textbox(fitz.Rect(45 + 298 , 790 + 1/fontsize * 60, 100 + 280, 830), buffer = text_output[7] + '€', fontname ="/Rosmatika",   align = 1, fontsize = 12+ np.log(max(int(text_output[7]),1)**2))

    fontsize = 12+ np.log(max(int(text_output[8]),1)**2)
    page.insert_textbox(fitz.Rect(492, 790+ 1/fontsize * 60, 550, 830), buffer = text_output[8] + '€', fontname ="/Rosmatika", align = 1, fontsize = 12+ np.log(max(int(text_output[8]),1)**2))
        
    # PAGE 2
    page = doc.load_page(1)
    page.insert_textbox(fitz.Rect(240, 65+10, 300, 100), buffer = text_output[0], fontname ="/Rosmatika", fontfile = currdir + "Fonts//Rosmatika (DEMO).ttf",  align = 1, fontsize = 20)
    page.insert_textbox(fitz.Rect(170, 65+55, 300, 150), buffer = text_output[1], fontname ="/Rosmatika",    align = 1, fontsize = 20)
    page.insert_textbox(fitz.Rect(190, 65+95, 300, 250), buffer = text_output[2], fontname ="/Rosmatika",    align = 1, fontsize = 20)
    page.insert_textbox(fitz.Rect(480, 90, 595, 110), buffer = text_output[3], fontname ="/Rosmatika",    align = 1, fontsize = 20)
    page.insert_textbox(fitz.Rect(520, 158, 595, 190), buffer = text_output[4], fontname ="/Rosmatika",    align = 1, fontsize = 20)
            
    page.insert_textbox(fitz.Rect(170, 285, 210, 340), buffer = str(round(df2temp['q2'][i],2)), fontname ="Anton",  fontfile = currdir + "Fonts//Anton-Regular.ttf",   align = 1, fontsize = 21.4)
          
    page.insert_textbox(fitz.Rect(500, 285, 540, 340), buffer = str(round(df2temp['q1'][i],2)), fontname ="/Anton",     align = 1, fontsize = 21.4)
          

        
        
    page.insert_image(fitz.Rect(-180,400,355,560), filename='f.png')
    page.insert_image(fitz.Rect(23,400,558,560), filename='f.png')
    page.insert_image(fitz.Rect(225,400,780,560), filename='f.png')

    page.insert_image(fitz.Rect(-180,515,355,675), filename='f.png')
    page.insert_image(fitz.Rect(23,515,558,675), filename='f.png')
    page.insert_image(fitz.Rect(225,515,780,675), filename='f.png')

    page.insert_image(fitz.Rect(-78.5,615,456.5,785), filename='f.png')
    page.insert_image(fitz.Rect(124.5,615,659.5,785), filename='f.png')

    page.insert_image(fitz.Rect(23,730,558,890), filename='f.png')

    # page.insert_image(fitz.Rect(0,0, 595,842), filename = "D:/Documents/boti4881/Downloads/red.png")



    # PAGE 2
    page = doc.load_page(2)

    fontsize = 12+ np.log(max(int(text_output[5]),1)**2)
    page.insert_textbox(fitz.Rect(40, 135 + 1/fontsize * 60, 100, 200), buffer = text_output[5]  + '€', fontname ="/Rosmatika", fontfile = currdir + "Fonts//Rosmatika (DEMO).ttf",  align = 1, fontsize = fontsize)

    fontsize = 12+ np.log(max(int(text_output[6]),1)**2)
    page.insert_textbox(fitz.Rect(40 + 135, 135+ 1/fontsize * 60 , 100 + 137, 200), buffer = text_output[6] + '€', fontname ="/Rosmatika",   align = 1, fontsize = 12+ np.log(max(int(text_output[6]),1)**2))
        
    fontsize = 12+ np.log(max(int(text_output[7]),1)**2)
    page.insert_textbox(fitz.Rect(40 + 298 , 135 + 1/fontsize * 60, 100 + 280, 200), buffer = text_output[7] + '€', fontname ="/Rosmatika",   align = 1, fontsize = 12+ np.log(max(int(text_output[7]),1)**2))

    fontsize = 12+ np.log(max(int(text_output[8]),1)**2)
    page.insert_textbox(fitz.Rect(485, 135+ 1/fontsize * 60, 550, 200), buffer = text_output[8] + '€', fontname ="/Rosmatika", align = 1, fontsize = 12+ np.log(max(int(text_output[8]),1)**2))
     

    page.insert_textbox(fitz.Rect(240, 165+70, 300, 300), buffer = text_output[0], fontname ="/Rosmatika",   align = 1, fontsize = 20)
    page.insert_textbox(fitz.Rect(170, 165+115, 300, 350), buffer = text_output[1], fontname ="/Rosmatika",    align = 1, fontsize = 20)
    page.insert_textbox(fitz.Rect(190, 165+155, 300, 350), buffer = text_output[2], fontname ="/Rosmatika",    align = 1, fontsize = 20)
    page.insert_textbox(fitz.Rect(480, 190+60, 595, 300), buffer = text_output[3], fontname ="/Rosmatika",    align = 1, fontsize = 20)
    page.insert_textbox(fitz.Rect(520, 318, 595, 400), buffer = text_output[4], fontname ="/Rosmatika",    align = 1, fontsize = 20)
       

    page.insert_image(fitz.Rect(10,400,390,700), filename='pn.png')

    page = doc.load_page(3)
    page.insert_image(fitz.Rect(-70,20,700,420), filename='s.png')
    page.insert_image(fitz.Rect(10,450,410,750), filename='ps.png')

    doc.save("Vin" + str(vin) +".pdf")
    doc.close()
