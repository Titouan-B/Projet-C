# -*- coding: utf-8 -*-
"""
Created on Fri Nov 25 22:12:03 2022

@author: boti4881
"""



# Slider année


import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
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

def ImportData(filename):
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

# =============================================================================

def ConsoPourcentage(df3_count, vin, text_output):
    
    countvin = df3_count.loc[df3_count["numvin"]== vin]
    total = int(countvin['q13_1'] + countvin['q13_2']+ countvin['q13_3'] + countvin['q13_4'] + countvin['q13_5'])
    
    for col in enumerate(countvin.keys()[1:6].tolist()):
        text_output.append("Situation "+str(col[0]+1)+" : " + str(float(countvin[col[1]]/total)))
    return text_output
# =============================================================================


def TranchePrix(df4transformed,vin,text_output):
    df4data = df4transformed.loc[df4transformed["numvin"] == vin]

    dixquinze = df4data['q14'].loc[df4data['q14'] == 1].count()
    quinzevingt = df4data['q14'].loc[df4data['q14'] == 0].count()
    vingtvingcinq = df4data['q14'].loc[df4data['q14'] == -1].count()
    vingcinqplus = df4data['q14'].loc[df4data['q14'] == 2].count()
    
    total = dixquinze + dixquinze + vingtvingcinq + vingcinqplus

    text_output += ['10-15€ : '+ str(dixquinze/total)]
    text_output += ['15-20€ : '+ str(quinzevingt/total)]
    text_output += ['20-25€ : '+ str(vingtvingcinq/total)]
    text_output += ['+25€ : '+ str(vingcinqplus/total)]
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
    label = ["Positif", "Neutre", "Négatif"]

#Compter le nombre de valeurs positives/neutres/négatives
    pos = winedata['q3'].loc[winedata['q3'] == 1].count()
    neu = winedata['q3'].loc[winedata['q3'] == 0].count()
    neg = winedata['q3'].loc[winedata['q3'] == -1].count()
    
    val = [pos,neu,neg]

# ajouter les données et la couleur
    label.append("")
    val.append(sum(val))
    colors = ['gold', 'grey', 'black', 'k']

# plot + titre
    fig = plt.figure(figsize=(8,6),dpi=100)
    wedges, labels = plt.pie(val, wedgeprops=dict(width=0.4,edgecolor='k'),labels=label, colors=colors)
    wedges[-1].set_visible(False)
    plt.title('Comment votre bouteille a affecté les consos', fontfamily = "Arial Rounded MT Bold")
    plt.savefig('f.png', format='png', transparent=True, dpi=100)
    return()

# =============================================================================



def GraphPositionVin(df2, vin):
# Graph du vin par rapport aux autres
    plt.clf() #Clear previous graphs in memory
    custom_params = {"axes.spines.right": False, "axes.spines.top": False}
    sns.set(style="ticks", rc=custom_params)
    # Mise en forme des données
    rdf1 = df2['q1'].reset_index()
    rdf2 = df2['q2'].reset_index()
    
    # Faire le plot + ligne du vin
    meanq1 = rdf1['q1'].loc[rdf1['numvin'] == vin]
    meanq2 = rdf2['q2'].loc[rdf2['numvin'] == vin]
    
    # Lignes insiquant la position du vin
    plt.axvline(float(meanq1), 0,0.95, color = 'gold',ls = '--')
    plt.text(float(meanq2)-0.4, 0.6,"Votre vin", fontfamily = "Arial Rounded MT Bold")
    plt.axvline(float(meanq2), 0,0.95, color = 'black', ls = '--')
    
    # Faire les deux plots + titre et enlever l'axe y
    fig2 = sns.kdeplot(rdf1['q1'], shade=True, bw_method=0.5, color="gold")
    fig2 = sns.kdeplot(rdf2['q2'], shade=True, bw_method=0.5, color="black")
    fig2.figure.suptitle('Votre vin par rapport aux autres !', fontfamily = "Arial Rounded MT Bold")
    fig2.get_yaxis().set_visible(False)
    plt.savefig('p.png', format='png', transparent=True, dpi=100)
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

def RadarplotVin(pnnmoy, vin):

    plt.clf() #Clear previous graphs in memory

# Recherche des points max/min
    questions = pnnmoy.keys()[1:4].values.tolist() #on prend les noms des colonnes

    maxval = []
    minval = []
    for i in range(len(questions)):
        maxval.append(max(pnnmoy[questions[i]]))
        minval.append(min(pnnmoy[questions[i]]))
        
        # Echelle du radarplot
        markers = [0, 1, 2, 3, 4, 5, 6]
        
        # Catégories présentes sur le graphique (exemple ici)
        categories = ["Finesse des bulles", "Tenue cordon", "Robe"]
        categories = [*categories, categories[0]]
        
        # Données du graphique (colonnes 4,5,6 pour l'exemple)
        data = pnnmoy.loc[pnnmoy["numvin"] == vin].values.tolist()[0][4:7]
        #les valeurs ici sont entre 0 et 1 donc je remap rapidement de 1 à 6
        data = list(map(lambda item: item * 6, data))
        data = [*data, data[0]]

# Points max et min - remap aussi
    minval = list(map(lambda item: item * 6, minval))
    maxval = list(map(lambda item: item * 6, maxval))
    minval = [*minval, minval[0]]
    maxval = [*maxval, maxval[0]]
    
    # Gestion des angles et plot
    label_loc = np.linspace(start=0, stop=2 * np.pi, num=len(data))
    plt.figure(figsize=(8, 8))
    ax = plt.subplot(polar=True) #le polar fait que c'est un graph radial
    
    # Supprimer lex axes pour mettre devant un template sur le pdf ?
    ax.axis('off')
    
    # Lignes + points du vin actuel
    plt.plot(label_loc, data,'o-', label=str(vin), color='gold',linewidth=4.0)
    plt.scatter(label_loc, data, color='gold',s=200)
    
    # Points max/min
    plt.scatter(label_loc, maxval, color='green',s=100)
    plt.scatter(label_loc, minval, color='red',s=100)
    
    #Plot + titre
    plt.yticks(markers)
    plt.title('Caractéristiques du vin '+ str(vin), size=20)
    lines, labels = plt.thetagrids(np.degrees(label_loc))#, labels=questions)
    plt.savefig('r.png', format='png', transparent=True, dpi=100)
    return()
    

# =============================================================================

# Modifier pdf dans Python


def CreaPDF(vin, text_per_pdf,filedir):
    former_dir = getcwd()
    chdir(filedir)
    doc = fitz.open()
    page = doc.new_page()
    for page in doc:
        page.clean_contents()   
    for text in enumerate(text_per_pdf):
        page.insert_textbox(fitz.Rect(0,210+float(text[0])*20,1000,250+float(text[0])*20), text[1])
        page.insert_image(fitz.Rect(0,0,200,200), filename=str(former_dir + '\\r.png'))
        page.insert_image(fitz.Rect(200,0,400,200), filename=str(former_dir + '\\p.png'))
        page.insert_image(fitz.Rect(400,0,600,200), filename=str(former_dir + '\\f.png'))
    doc.save("Vin" + str(vin) +".pdf")
    doc.close()
