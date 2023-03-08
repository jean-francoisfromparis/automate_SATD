import csv
import os
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path
from tkinter import *
from tkinter import filedialog, messagebox, messagebox as msg, ttk
from tkinter.messagebox import showinfo
from tkinter.ttk import Progressbar

import numpy as np
import pandas as pd
import pyexcel_ods3 as pe
from PIL import Image, ImageTk
from pandastable import Table
from pyexcel_ods import save_data
from pynput.keyboard import Controller
from selenium import webdriver
from selenium.common import TimeoutException, WebDriverException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

keyboard = Controller()


def __init__(self, progress):
    self.progress = progress
    global delay
    delay = 3


# Fonction pour retrouver le chemin d'accès
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


# Procédure pour la consultation des oppositions
def main():
    # Délai entre opérations automate. Pour des numéros non entiers il faut utiliser le point pas la virgule
    delay = 1

    # ########################################

    # ##Saisie du nom utilisateur et mot de passe
    login = EnterTable4.get()
    mot_de_passe = EnterTable5.get()

    ## Saisie de numéro de dossier:
    numeroDossier = EnterTable6.get()

    wd_options = Options()

    wd_options.headless = True
    wd_options.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'
    wd_options.set_preference('detach', True)
    wd = webdriver.Firefox(options=wd_options)
    try:
        wd.get(
            'http://medoc.ia.dgfip:8141/medocweb/presentation/md2oagt/ouverturesessionagent/ecran/ecOuvertureSessionAgent'
            '.jsf')
    except WebDriverException:
        messagebox.showinfo("Service Interrompu !", "Le service est indisponible\n pour l'instant")
        wd.close()

    ## Saisir utilisateur
    ##Saisir utilisateur
    time.sleep(delay)
    # script = f'''identifant = document.getElementById('identifiant'); identifiant.setAttribute('type','hidden'); identifiant.setAttribute('value',"{login}");'''
    script = f'''identifant = document.getElementById('identifiant'); identifiant.setAttribute('type','hidden'); identifiant.setAttribute('value',"youssef.atigui");'''
    wd.execute_script(script)

    ## Saisie mot de pass
    time.sleep(delay)
    # wd.find_element(By.ID, 'secret_tmp').send_keys(mot_de_passe)
    wd.find_element(By.ID, 'secret_tmp').send_keys("1")

    time.sleep(delay)
    wd.find_element(By.ID, 'secret_tmp').send_keys(Keys.RETURN)

    try:
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'ligneServiceHabilitation')))
    except TimeoutException:
        messagebox.showinfo("Service Interrompu !", "Le service est indisponible\n pour l'instant")
        wd.close()

    ## Saisir service
    wd.find_element(By.ID, 'nomServiceChoisi').send_keys('0070100')
    time.sleep(delay)
    wd.find_element(By.ID, 'nomServiceChoisi').send_keys(Keys.TAB)

    ## Saisir habilitation
    time.sleep(delay)
    wd.find_element(By.ID, 'habilitation').send_keys('1')
    time.sleep(delay)
    wd.find_element(By.ID, 'habilitation').send_keys(Keys.ENTER)

    ## Boucle sur le fichier selon le nombre de lignes indiquées

    ## Création d'un Redevable
    ## Arriver à la transaction 3-17

    time.sleep(delay)
    wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys('3')
    wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys(Keys.ENTER)
    time.sleep(delay)
    WebDriverWait(wd, 20).until(
        EC.presence_of_element_located((By.ID, 'bmenuxtableMenus:16:outputBmenuxBrmenx04LibelleLigneProposee')))
    wd.find_element(By.ID, 'bmenuxtableMenus:16:outputBmenuxBrmenx04LibelleLigneProposee').click()

    ## Saisie numéro de Dossier
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))

    wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(numeroDossier)
    wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(Keys.ENTER)

    # try:
    #     WebDriverWait(wd, 20).until(EC.text_to_be_present_in_element((By.CLASS_NAME, 'ui-messages-error-summary'),
    #                                                                  'DOSSIER DEJA UTILISE PAR UN AUTRE POSTE  - '
    #                                                                  'ATTENTE OU ABANDON -                 '))
    #     showinfo("Affichage opposition", "Le dossier de l'opposant " + numeroDossier + " est déjà utilisé et doit être "
    #                                                                                    "purgé ")
    #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
    #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
    #     wd.quit()
    # except FileNotFoundError as e:
    #     print(e)
    #     msg.showerror('Error in opening file', e)
    # finally:
    #     pass
    # if EC.text_to_be_present_in_element((By.CLASS_NAME, 'ui-messages-error-summary'),
    #                                                                 'DOSSIER DEJA UTILISE PAR UN AUTRE POSTE  - '
    #                                                                 'ATTENTE OU ABANDON -                 '):
    #     showinfo("Affichage opposition", "Le dossier de l'opposant " + numeroDossier + " est déjà utilisé et doit être "
    #                                                                                    "purgé ")
    # else:
    #     pass

    ## Saisie du choix Lister
    time.sleep(delay)
    time.sleep(delay)
    time.sleep(delay)
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI')))

    wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys('I')
    wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys(Keys.TAB)

    ## Récupération du tableau des oppositions
    time.sleep(delay)
    time.sleep(delay)
    time.sleep(delay)

    currentUrl = wd.current_url
    compteur = 0
    k = 0
    while True:
        if currentUrl == 'http://medoc.ia.dgfip:8141/medocweb/presentation/transactions/redevable/pa33g/ecran' \
                         '/Pa33GTx317.jsf':
            time.sleep(delay)
            time.sleep(delay)
            time.sleep(delay)
            try:
                globals()[f"webtable_df{k}"] = pd.read_html(
                    wd.find_element(By.XPATH, '//*[@id="b33GlistLigneOperationPanel"]').get_attribute('outerHTML'))[
                    1]
                webtable_df1 = \
                    pd.read_html(
                        wd.find_element(By.XPATH, '//*[@id="b33GlistLigneOperationPanel"]').get_attribute('outerHTML'))[
                        1]
                time.sleep(delay)
                time.sleep(delay)
                time.sleep(delay)
                # dataTable = pd.concat([globals()[f"webtable_df{k}"]], ignore_index=True)
                time.sleep(delay)
                time.sleep(delay)
                time.sleep(delay)
                # print("dataTable au tour" + str(k) + ": \n", globals()[f"webtable_df{k}"])
                last_index = globals()[f"webtable_df{k}"]['Unnamed: 0']
                index_list = globals()[f"webtable_df{k}"]['Unnamed: 0'].isnull().values.any()
                if index_list:
                    print(index_list)
                    print(last_index)
                    print(globals()[f"webtable_df{k}"].dtypes)
                    print("fin: " + str(compteur))
                    time.sleep(delay)
                    time.sleep(delay)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    dataTable = pd.DataFrame()
                    for k in range(0, compteur + 1):
                        time.sleep(delay)
                        time.sleep(delay)
                        dataTable = dataTable.append([globals()[f"webtable_df{k}"]], ignore_index=True)
                        # print("dataTable au tour" + str(k + 1) + ": \n", globals()[f"webtable_df{k}"])
                        time.sleep(delay)
                        time.sleep(delay)
                    print("dataTable à la fin: \n", dataTable)
                    indice = pd.to_numeric(dataTable['Unnamed: 0']).fillna(0).astype(int)
                    FRP = pd.to_numeric(dataTable['Unnamed: 1']).fillna(0).astype(int)
                    name = dataTable['Unnamed: 2']
                    credit = pd.to_numeric(dataTable['Unnamed: 3']).fillna(0)
                    montant = dataTable['Unnamed: 10']
                    levee = dataTable['Unnamed: 16']
                    fields = {'id': indice, 'FRP': FRP, 'DENOMINATION': name, ' CREDIT D\'IMPOT': credit,
                              'Montant': montant,
                              'LEVEE': levee}
                    table = pd.DataFrame(fields)
                    filename = EnterTable6.get() + '_liste_créances_' + datetime.now().strftime(
                        '%Y-%m-%d-%H-%M-%S') + '.csv'
                    table.to_csv(filename, columns=fields, index=FALSE)

                    try:
                        time.sleep(delay)
                        time.sleep(delay)
                        time.sleep(delay)
                        liste = csv.reader(open(filename), delimiter=',')
                        tabControl.add(tab3, text='liste des oppositions')
                        table1 = Table(tab3, dataframe=table, read_only=True, index=FALSE)
                        table1.place(y=120)
                        table1.show()

                    except FileNotFoundError as e:
                        print(e)
                        msg.showerror('Error in opening file', e)
                    ## Validation de la sortie du formulaire
                    time.sleep(delay)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                else:
                    # time.sleep(delay)
                    time.sleep(delay)
                    time.sleep(delay)
                    if not globals()[f"webtable_df{k}"].empty:
                        dataTable = pd.concat([globals()[f"webtable_df{k}"]], ignore_index=True)
                        wd.find_element(By.ID, 'inputB33gnaviY33GnavichChoixSurB33Gnavi').send_keys('S')
                        wd.find_element(By.ID, 'inputB33gnaviY33GnavichChoixSurB33Gnavi').send_keys(Keys.ENTER)
                        print("Avant: " + str(compteur))
                        compteur += 1
                        k += 1
                        print("après: " + str(compteur))
                    else:
                        WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                        wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                        wd.quit()
                        showinfo("Affichage opposition",
                                 "L'opposant " + numeroDossier + " n'a pas d'opposition en cours ")
            except:
                WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                wd.quit()
                showinfo("Affichage opposition", "L'opposant " + numeroDossier + " n'a pas d'opposition en cours ")
                break

    # finally:
    #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
    #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
    #     wd.quit()
    #     showinfo("Affichage opposition", "L'opposant " + numeroDossier + " n'a pas d'opposition en cours ")


def create_opposition(headless):
    delay = 3

    # Etablissement du progressBar

    pb = progressbar(tab6)
    progressbar_label = Label(tab6, text=f"Le travail commence. L'automate se connecte...")
    label_y = 390
    progressbar_label.place(x=250, y=label_y)
    tab6.update()

    time.sleep(delay)

    ##Prend la ligne du fichier depuis laquelle commencer à lire
    # while True:
    #     line = EnterTable2.get()
    #     if line.isnumeric():  ##vérifie que ça soit un numéro
    #         line = int(line)  ##ajuste l'indice
    #         break
    #     else:
    #         messagebox.showerror("Erreur de saisie", 'Saisie incorrecte, réessayez')
    #         exit()

    ##Combien de lignes du fichier traiter
    # while True:
    #     line_amount = EnterTable3.get()
    #     if line_amount.isnumeric():
    #         line_amount = int(line_amount)
    #         break
    #     else:
    #         messagebox.showerror("Erreur de saisie", 'Saisie incorrecte, réessayez')
    #         exit()

    ## Prend les données depuis le fichier, crée une liste de listes (ou "array"), oú chaque liste est
    ## une ligne du fichier Calc. Il faut faire ça parce que pyxcel_ods prend les données sous forme
    ## de dictionnaire.
    donnees_creation_opposition = pe.get_data(File_path)
    source_rep = os.getcwd()
    filename1 = 'donnees_creation_opposition_sortie' + datetime.now().strftime('_%Y-%m-%d') + '.ods'
    filepath1 = source_rep + '/donnees_sortie/donnees_sortie' + datetime.now().strftime('_%Y-%m-%d') + '/' + filename1
    # df = pd.read_excel(File_path)
    print("filepath1: \n", filepath1)
    print("----------------------------------------------------------------------------")
    match os.path.isfile(filepath1):
        case True:
            donnees_creation_opposition_sortie = pd.read_excel(filepath1)
            donnees_creation_opposition = pd.read_excel(File_path)
            donnees_creation_opposition["Numéro d'Opération"] = ""
            donnees_creation_opposition["Date d'exécution"] = ""
            donnees_creation_opposition["Dossiers traités"] = ""
            donnees_creation_opposition["Dossiers traités"] = ""
            print("dataframe des données d'entrée : \n", donnees_creation_opposition)
            print("----------------------------------------------------------------------------")
            # Enlever les données manquantes du fichier d'entrée
            taille_donnee_entree = donnees_creation_opposition.shape[0]
            print("Taille des données d'entrées :  \n", taille_donnee_entree)
            print("----------------------------------------------------------------------------")
            print("les données : \n", donnees_creation_opposition)
            print("----------------------------------------------------------------------------")
            for i in range(taille_donnee_entree):
                if donnees_creation_opposition.loc[i].isnull().any() or \
                        donnees_creation_opposition["Date d’effet = date réception SATD"].loc[i] == 'NaT':
                    donnees_creation_opposition.drop(i, inplace=True)
            taille_donnee_entree1 = donnees_creation_opposition.shape[0]
            print("Taille des données d'entrées après suppression lignes incomplétes : \n", taille_donnee_entree1)
            print("----------------------------------------------------------------------------")
            print("les données après suppression des lignes incomplétes :  \n", donnees_creation_opposition)
            print("----------------------------------------------------------------------------")
            # Enlever les données déjà passées du fichier d'entrée
            old_data_done = donnees_creation_opposition_sortie[
                (donnees_creation_opposition_sortie['Dossiers traités'] == 'X')]
            old_data_done_list = old_data_done["Réf jugement validité = réf SATD"]
            print("liste des données déja passéés \n", old_data_done_list)
            for element in old_data_done["Réf jugement validité = réf SATD"]:
                old_data_done_list_index = donnees_creation_opposition[
                    donnees_creation_opposition["Réf jugement validité = réf SATD"] == element].index
                donnees_creation_opposition.drop(old_data_done_list_index, inplace=True)
                print("Dataframe après suppression des données déjà enregistré : ", donnees_creation_opposition)
            data = donnees_creation_opposition.values.tolist()
            old_data = old_data_done.values.tolist()
            nb_ligne = len(data)
            print("nb ligne sortie 1: \n", nb_ligne)
            print("Les données initiales à ne pas utiliser: \n", old_data)
            print("Les données initiales: \n", data)

            # join = pd.concat([donnees_creation_opposition, donnees_creation_opposition_sortie])
            # print("la fusion : \n", join["Réf jugement validité = réf SATD"])
            # print("----------------------------------------------------------------------------")
            # old_data = donnees_creation_opposition_sortie[(donnees_creation_opposition_sortie['Dossiers traités'] == 'X') | (
            #         donnees_creation_opposition_sortie['Données manquantes'] == '∅')].values.tolist()
            # all_old_data_list = donnees_creation_opposition_sortie["Réf jugement validité = réf SATD"].values.tolist()
            # print("liste des données", old_data)
            # print("----------------------------------------------------------------------------")
            # result = join[-join["Réf jugement validité = réf SATD"].isin(all_old_data_list)]
            # print("résultat de la jointure : ",result)
            # print("----------------------------------------------------------------------------")
            # data = donnees_creation_opposition_sortie[(donnees_creation_opposition_sortie['Dossiers traités'] != 'X') & (
            #         donnees_creation_opposition_sortie['Données manquantes'] != '∅')]
            # data = pd.concat([data, result], ignore_index=True).values.tolist()
            # nb_ligne = len(data)
            # print("nb ligne sortie 1: ", nb_ligne)
            # print("Les données initiales à ne pas utiliser: ", old_data)
            # print("Les données initiales: ", data)
        case False:
            donnees_creation_opposition_sortie = pe.get_data(File_path)
            print("Mauvaise sortie")
            donnees_creation_opposition_sortie['Feuille1'][0].append("Numéro d'Opération")
            donnees_creation_opposition_sortie['Feuille1'][0].append("Date d'exécution")
            donnees_creation_opposition_sortie['Feuille1'][0].append("Dossiers traités")
            # donnees_creation_opposition_sortie['Feuille1'][0].append("Dossiers traités")
            donnees_creation_opposition = pd.read_excel(File_path)
            donnees_creation_opposition["Date d’effet = date réception SATD"] = donnees_creation_opposition[
                "Date d’effet = date réception SATD"].astype(str)

            nb_ligne = donnees_creation_opposition.shape[0]
            ligne_incomplete = list()
            for i in range(nb_ligne):
                if donnees_creation_opposition.loc[i].isnull().any() or \
                        donnees_creation_opposition["Date d’effet = date réception SATD"].loc[i] == 'NaT':
                    ligne_incomplete.append('∅')
                    # print(ligne_incomplete)
                else:
                    ligne_incomplete.append('')
                    # print(ligne_incomplete)
            donnees_creation_opposition["Dossiers traités"] = ligne_incomplete
            print(ligne_incomplete)
            # donnees_creation_opposition.insert(loc=5, column='Lignes non traitées', value=ligne_incomplete)
            # print("type de données: ", donnees_creation_opposition)
            old_data = donnees_creation_opposition[donnees_creation_opposition["Dossiers traités"] == '∅'].values \
                .tolist()
            print("les données non gardé ligne 346 \n", old_data)
            data = donnees_creation_opposition[donnees_creation_opposition["Dossiers traités"] != '∅'].values.tolist()
            print("les données d'entrée ligne 347 \n", data)
            nb_ligne = len(data)
            print(nb_ligne)
    # exit()
    print("les données d'entrée ligne 373 \n", data)
    # df = pd.DataFrame(
    #     columns=["Indice", "FRP société", "FRP opposant", "Montant", "Date d’effet = date réception SATD",
    #              "Numéro d'Opération", "Date d'exécution", "Dossiers traités"])
    # Condition qui vérifie que chaque cellule de la colonne rib, à part le header, est vide, d'après le besoin case
    # vide = rang 1, si l'item correspondant au rang est vide il prend la valeur "1" utilisable dans la boucle
    # d'automatisation. Cette condition sert à s'assurer que l'on aura une valeur pour le rang, s'il n'y a pas de
    # valeur la liste est vide et ça génère une erreur taille_data donne le nombre d'items+1 dans le dico,
    # puisque python boucle à partir de 0, dans notre cas, c'est le nombre de listes qui est de 11 (10 + liste
    # headers) C'est pour cela que je boucle de 0 à taille_data - 2 pour ne pas inclure la liste des headers.
    # taille_data = len(data)
    # print("taille_data : ", len(data))
    # last_item_index0 = len(data[0]) - 1
    # print("last_item_index0 : ", last_item_index0)
    # last_item_index1 = len(data[1]) - 1
    # for i in range(taille_data - 2):
    #     if last_item_index0 != len(data[i + 1]) - 1:
    #         data[i + 1].append(str(1))
    #########################################

    ## Saisie du nom utilisateur et mot de passe
    login = EnterTable4.get()
    mot_de_passe = EnterTable5.get()

    ## Saisie de numéro de dossier :
    # numeroDossier = EnterTable6.get()

    ## Saisie de la référence de jugement :
    # reference_de_jugement = EnterTable10.get()

    wd_options = Options()
    wd_options.headless = headless
    wd_options.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'
    wd_options.set_preference('detach', True)
    wd = webdriver.Firefox(options=wd_options)
    # wd = webdriver.Firefox(executable_path=GeckoDriverManager().install(), options=wd_options)
    ## TODO Passer au service object
    # wd.get(
    #     'https://portailmetierpriv.ira.appli.impots/cas/login?service=http%3A%2F%2Fmedoc.ia.dgfip%3A8141%2Fmedocweb'
    #     '%2Fcas%2Fvalidation')  # adresse MEDOC DGE

    wd.get(
        'http://medoc.ia.dgfip:8121/medocweb/presentation/md2oagt/ouverturesessionagent/ecran'
        '/ecOuvertureSessionAgent.jsf')  # adresse MEDOC Classic
    ##Saisir utilisateur
    time.sleep(delay)
    # script = f'''identifant = document.getElementById('identifiant'); identifiant.setAttribute('type','hidden'); identifiant.setAttribute('value',"{login}");'''
    script = f'''identifant = document.getElementById('identifiant'); identifiant.setAttribute('type','hidden'); 
    identifiant.setAttribute('value',"youssef.atigui"); '''
    wd.execute_script(script)

    ## Saisie mot de pass
    time.sleep(delay)
    # wd.find_element(By.ID, 'secret_tmp').send_keys(mot_de_passe)
    wd.find_element(By.ID, 'secret_tmp').send_keys("1")

    time.sleep(delay)
    wd.find_element(By.ID, 'secret_tmp').send_keys(Keys.RETURN)
    try:
        WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'ligneServiceHabilitation')))
    except TimeoutException:
        messagebox.showinfo("Service Interrompu !", "Le service est indisponible\n pour l'instant")
        wd.close()
    ## Saisir service
    # wd.find_element(By.ID, 'nomServiceChoisi').send_keys('0070100')  # FRP MEDOC DGE
    wd.find_element(By.ID, 'nomServiceChoisi').send_keys('6200100')
    time.sleep(delay)
    wd.find_element(By.ID, 'nomServiceChoisi').send_keys(Keys.TAB)

    ## Saisir habilitation
    try:
        time.sleep(delay)
        wd.find_element(By.ID, 'habilitation').send_keys('1')
        time.sleep(delay)
        wd.find_element(By.ID, 'habilitation').send_keys(Keys.ENTER)
    except:
        progressbar_label.destroy()
        WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
        messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
        messagebox.showinfo("Service Interrompu !", messages)
        wd.close()

    time.sleep(delay)
    wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys('3')
    wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys(Keys.ENTER)
    time.sleep(delay)

    ## Création d'un Redevable
    ## Arriver à la transactionv 3-17
    try:
        WebDriverWait(wd, 20).until(
            EC.presence_of_element_located((By.ID, 'bmenuxtableMenus:16:outputBmenuxBrmenx04LibelleLigneProposee')))
        wd.find_element(By.ID, 'bmenuxtableMenus:16:outputBmenuxBrmenx04LibelleLigneProposee').click()
    except:
        progressbar_label.destroy()
        messagebox.showinfo("Service Interrompu !", "La transaction création des oppositions ne semblent pas être "
                                                    "disponible. Veuillez tester manuellement avant de redémarrer "
                                                    "l'automate.")
        WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        wd.close()

    progressbar_label.destroy()

    ## Boucle sur le fichier selon le nombre de lignes indiquées
    # for j in range(nb_ligne):
    j = 0
    while True:
        print("N° de ligne à la ligne 510: ", j)
        source_rep = os.getcwd()
        destination_rep = source_rep + '/archive_SATD/archive' + datetime.now().strftime('_%Y-%m-%d')
        num_of_secs = 60
        m, s = divmod(num_of_secs * (nb_ligne + 1), 60)
        min_sec_format = '{:02d}:{:02d}'.format(m, s)
        progressbar_label = Label(tab6,
                                  text=f"Le travail est en cours: {pb['value']:.2f}%  ~  il reste environ {min_sec_format}")
        progressbar_label.place(x=250, y=label_y)
        tab6.update()

        ## Saisie numéro de Dossier
        # while True:
        # try:
        #     print("numero de dossier : ", data[i][0])
        #     WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))
        #     wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[i][0])
        #     wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(Keys.ENTER)
        # except:
        #     print("erreur")
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
        #
        #     errorMessages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
        #     messages = f"{errorMessages} + \n Le dossier N°{data[i][0]} est ouvert par un autre agent ou verrouillé." \
        #                f"\n Vous pouvez relancer le processus. Cette ligne sera exclu et pourra être relancer dans " \
        #                f"45 minutes"
        #     messagebox.showinfo("Service Interrompu !", messages)
        #     time.sleep(delay)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        while True:
            print("le N° de ligne est à la ligne 547 :", j)
            print("numéro de dossier : ", data[j][0])
            WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))
            wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[j][0])
            wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(Keys.ENTER)
            errorMessages = ""
            print("messages d'erreur: ", errorMessages)
            time.sleep(delay)
            time.sleep(delay)
            time.sleep(delay)
            WebDriverWait(wd, 200).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            errorMessages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            print("messages d'erreur: ", errorMessages)
            messageDossierVerrouille = "DOSSIER DEJA UTILISE PAR UN AUTRE POSTE  - ATTENTE OU ABANDON - ".replace(" ",
                                                                                                                  "")

            errorMessagesIsPresent = errorMessages.replace(" ", "") == messageDossierVerrouille
            time.sleep(delay)
            time.sleep(delay)
            if errorMessagesIsPresent:
                messages = f"{errorMessages} \n Le dossier N°{data[j][0]} est ouvert par un autre agent ou verrouillé." \
                           f"\n Vous pouvez relancer le processus. Cette ligne sera exclu et pourra être relancer dans " \
                           f"45 minutes"
                messagebox.showinfo("Dossier verrouillé !", messages)
                data[j].append('')
                data[j].append('\U0001F512')
                time.sleep(delay)
                # WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                # wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                j = j + 1
            else:
                break
            print("le N° de ligne est à la ligne 579 :", j)
            # exit()

        ## Saisie du choix Créer
        try:
            time.sleep(delay)
            time.sleep(delay)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI')))
            wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys('C')
            wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys(Keys.TAB)
            # print("ligne 473: ok")
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            # print("ligne 477")
            errorMessages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messages = errorMessages
            print(messages)
            time.sleep(delay)
            time.sleep(delay)
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Saisie du numéro de dossier créancier
        try:
            time.sleep(delay)
            time.sleep(delay)
            WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))
            # wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(numero_creancier_opposant)
            wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[j][1])
            wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(Keys.TAB)
            print("ligne 604: ok")
            print("le N° de ligne est à la ligne 605 :", j)  # print(data[i][1])
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            errorMessages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messages = errorMessages + "\nLa qualité de la connexion ne permet pas un bon fonctionnement de " \
                                       "l'automate. Veuillez essayer ultérieurement ! "
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Saisie de la suite
        try:
            time.sleep(delay)
            time.sleep(delay)
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33gsuitYa33G002ReponseSuite')))
            wd.find_element(By.ID, 'inputB33gsuitYa33G002ReponseSuite').send_keys('S')
            wd.find_element(By.ID, 'inputB33gsuitYa33G002ReponseSuite').send_keys(Keys.TAB)
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## SAISIE DES REFERENCES DE L'OPPOSITION
        ## Transport de créance
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance').send_keys('N')
            wd.find_element(By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance').send_keys(Keys.TAB)
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Saisie ATD
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GadtAdt')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GadtAdt').send_keys('O')
            wd.find_element(By.ID, 'inputB33ginf2Ya33GadtAdt').send_keys(Keys.TAB)
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Saisie du crédit
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GcredCreditIs')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GcredCreditIs').send_keys('N')
            wd.find_element(By.ID, 'inputB33ginf2Ya33GcredCreditIs').send_keys(Keys.TAB)
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Saisie Empêchement
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GempEmpechement')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GempEmpechement').send_keys('N')
            wd.find_element(By.ID, 'inputB33ginf2Ya33GempEmpechement').send_keys(Keys.TAB)
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Saisie Montant
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GmtMontant')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GmtMontant').send_keys(data[j][2])
            wd.find_element(By.ID, 'inputB33ginf2Ya33GmtMontant').send_keys(Keys.TAB)
            # print(data[i][2])
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Saisie de la Date d'Effet
        print(type(data[j][3]))
        if isinstance(data[j][3], str):
            date_d_effet = datetime.strptime(data[j][3], "%Y-%m-%d")
            print("ici c'est un string")
            print(date_d_effet.day)
        else:
            date_d_effet = data[j][3]
            print("ici ce n'est pas un string")
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour').send_keys(date_d_effet.day)
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour').send_keys(Keys.TAB)
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Saisie du Mois d'Effet
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois').send_keys(date_d_effet.month)
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois').send_keys(Keys.TAB)
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Saisie de l'Année d'Effet
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee').send_keys(date_d_effet.year)
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee').send_keys(Keys.TAB)
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Saisie de la référence de jugement
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite').send_keys(data[j][4])
            wd.find_element(By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite').send_keys(Keys.TAB)
            # print(data[i][4])
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Saisie de la date d'exécution de jugement
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementJour')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementJour').send_keys(Keys.TAB)

            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementMois')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementMois').send_keys(Keys.TAB)

            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementAnnee')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementAnnee').send_keys(Keys.TAB)
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Saisie de la date de renouvellement
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementJour')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementJour').send_keys(Keys.TAB)

            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementMois')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementMois').send_keys(Keys.TAB)

            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee').send_keys(Keys.TAB)
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Validation de la non saisie des dates
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee')))
            wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').send_keys('O')
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee').send_keys(Keys.TAB)
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Validation de la suite
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec')))
            wd.find_element(By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec').send_keys('S')
            wd.find_element(By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec').send_keys(Keys.TAB)
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Validation de la saisie de l'opposition
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33gvlcrYa33GvalcValidationCreation')))
            wd.find_element(By.ID, 'inputB33gvlcrYa33GvalcValidationCreation').send_keys('O')
            wd.find_element(By.ID, 'inputB33gvlcrYa33GvalcValidationCreation').send_keys(Keys.TAB)
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Capture numéro d'opération
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'outputB33gnopeYa33GnopeNOpe')))
            numero_ope = wd.find_element(By.ID, 'outputB33gnopeYa33GnopeNOpe').text
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Saisie de la fin de saisie
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33gnouvYa33GnvopNouvelleOpposition')))
            wd.find_element(By.ID, 'inputB33gnouvYa33GnvopNouvelleOpposition').send_keys('N')
            wd.find_element(By.ID, 'inputB33gnouvYa33GnvopNouvelleOpposition').send_keys(Keys.TAB)
        except:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        ## Validation de la sortie du formulaire
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        # except:
        #     progressbar_label.destroy()
        #     WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
        #     messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
        #     messagebox.showinfo("Service Interrompu !", messages)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.close()

        ## Marquage tâche faîte dans le fichier
        match os.path.isfile(filepath1):
            case True:
                data[j][5] = numero_ope
                data[j][6] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                data[j][7] = 'X'
                print("inscription des données dans la liste ligne 929", data)
            case False:
                data[j][3] = str(date_d_effet.strftime('%Y-%m-%d'))
                data[j].insert(5, numero_ope)
                data[j].insert(6, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
                data[j][7] = 'X'
                print("inscription des données ligne 936", data)
        print("le N° de ligne est  à la ligne 937:", j)

        ## Incrementation ProgressBar

        pb['value'] += 90 / nb_ligne
        progressbar_label.destroy()
        tab6.update()
        progress = pb['value']
        progressbar_label = Label(tab6,
                                  text=f"Le travail est en cours : {pb['value']:.2f}% il reste environ {min_sec_format}")
        progressbar_label.place(x=250, y=label_y)
        pb.update()
        tab6.update()
        print("le N° de ligne est  à la ligne 950:", j)
        if j < nb_ligne - 1:
            j += 1
        else:
            break
    columns = ["FRP société", "FRP opposant", "Montant", "Date d’effet = date réception SATD",
               "Réf jugement validité = réf SATD", "Numéro d'Opération", "Date d'exécution",
               "Dossiers traités"]
    data.insert(0, columns)

    print("les nouvelles data : \n", data)
    # source_rep = os.getcwd()
    destination_rep1 = source_rep + '/donnees_sortie/donnees_sortie' + datetime.now().strftime('_%Y-%m-%d')
    saved_file = destination_rep1 + '/' + filename1
    if not os.path.exists(destination_rep1):
        os.makedirs(destination_rep1)
    if os.path.exists(destination_rep1 + '/' + filename1):
        os.remove(destination_rep1 + '/' + filename1)
        for i in range(len(old_data)):
            old_data[i].insert(5, '')
            old_data[i].insert(6, '')
        print("old_data : \n", old_data)
        del data[0]
        print("data sans les entêtes (ligne 978)", data)
        # if old_data == [""]:

        numpyData = np.append(data, old_data, axis=0)
        data = list(numpyData)
        data.insert(0, columns)
        print("listData : \n", data)
        wd.close()
    else:
        for i in range(len(old_data)):
            if old_data[i][5] == '' & old_data[i][6] == '':
                del old_data[i][5]
                del old_data[i][6]
        print("old_data : \n", old_data)
        del data[0]
        print("data sans les entêtes (ligne 993)", data)
        if old_data == []:
            data = data
        else:
            numpyData = np.append(data, old_data, axis=0)
            data = list(numpyData)
        data.insert(0, columns)
        print("listData : \n", data)
        wd.close()

    save_data(saved_file, data)

    frp_opposant = list(zip(data[1]))
    # zipped = list(zip(data))
    # print("zipped", zipped)
    # data_df = pd.DataFrame.columns(
    #     ["Indice", "FRP société", "FRP opposant", "Montant", "Date d’effet = date réception SATD",
    #      "Numéro d'Opération", "Dossiers traités"])
    data_df = pd.DataFrame(data)

    print("le dataframe : ", data_df)

    try:
        time.sleep(delay)
        time.sleep(delay)
        time.sleep(delay)
        tabControl.add(tab4, text='liste des oppositions')
        table1 = Table(tab4, dataframe=data_df, read_only=True, index=FALSE)
        table1.place(y=120)
        table1.autoResizeColumns()
        table1.show()

    except FileNotFoundError as e:
        print(e)
        messagebox.showerror('Erreur de tableau', 'Il n\'y a pas de tableau à afficher')
    progressbar_label.destroy()
    tab2.update()
    progressbar_label = Label(tab2,
                              text=f"Le travail est maintenant fini! A bientôt")
    progressbar_label.place(x=250, y=label_y)
    wd.quit()


# Procédure de purge
def purge():
    # Délai entre opérations automate. Pour des numéros non entiers il faut utiliser le point pas la virgule
    delay = 1

    # ##Saisie du nom utilisateur et mot de passe
    login = EnterTable4.get()
    mot_de_passe = EnterTable5.get()

    ## Saisie de numéro de dossier:
    numeroDossier = EnterTable6.get()

    wd_options = Options()

    # wd_options.headless = True

    # wd_options.set_preference('detach', True)
    # wd = webdriver.Firefox(executable_path=GeckoDriverManager().install(), options=wd_options)
    wd_options.headless = False
    wd_options.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'
    wd_options.set_preference('detach', True)
    try:
        wd = webdriver.Firefox(options=wd_options)
    except WebDriverException:
        messagebox.showinfo("Service Interrompu !", "Votre système rencontre des difficultés à afficher le navigateur")

    url = 'http://media.ira.appli.impots/mediamapi/index.xhtml'
    try:
        wd.get(url)

    except WebDriverException:
        messagebox.showinfo("Service Interrompu !", "Le service est indisponible\n pour l'instant")
        wd.close()

    ## Saisir utilisateur
    ##Saisir utilisateur
    time.sleep(delay)
    # script = f'''identifant = document.getElementById('identifiant'); identifiant.setAttribute('type','hidden');
    # identifiant.setAttribute('value',"{login}");'''
    try:
        script = f'''identifant = document.getElementById('identifiant'); identifiant.setAttribute('type','hidden'); 
        identifiant.setAttribute('value',"youssef.atigui");'''
        wd.execute_script(script)
    except WebDriverException:
        messagebox.showinfo("Service Interrompu !", "Votre système rencontre des difficultés à afficher le navigateur")
        wd.close()
    ## Saisie mot de pass
    time.sleep(delay)
    # wd.find_element(By.ID, 'secret_tmp').send_keys(mot_de_passe)
    wd.find_element(By.ID, 'secret_tmp').send_keys("1")

    time.sleep(delay)
    wd.find_element(By.ID, 'secret_tmp').send_keys(Keys.RETURN)

    # try:
    #     WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'ligneServiceHabilitation')))
    # except TimeoutException:
    #     messagebox.showinfo("Service Interrompu !", "Le service est indisponible\n pour l'instant")
    #     wd.close()
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'j_idt146:j_idt147')))
    menu_list = wd.find_element(By.ID, "j_idt146:j_idt147")
    a = ActionChains(wd, 100)
    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="j_idt146:j_idt147"]/ul/li[1]')))
    service = wd.find_element(By.XPATH, '//*[@id="j_idt146:j_idt147"]/ul/li[1]')
    a.move_to_element(service).perform()
    time.sleep(1)
    purge_link_script = "document.evaluate('//*[@id=\"j_idt146:j_idt147\"]/ul/li[1]/ul/li[4]/a',document,null," \
                        "XPathResult.FIRST_ORDERED_NODE_TYPE,null,).singleNodeValue.click()"
    wd.execute_script(purge_link_script)

    try:
        WebDriverWait(wd, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, 'zone_validation')))
        purge_input = wd.find_element(By.CLASS_NAME, 'zone_validation')
        purge_input.click()
        nombre_de_verrou = wd.find_element(By.ID, 'formPurgerVerrouResultat:verrouPurger').text

        messagebox.showinfo("Purge", f"La purge a été opérée. \n {nombre_de_verrou} \n Vous pouvez "
                                     f"reprendre la création des oppositions. ")
        wd.close()
    except WebDriverException:
        messagebox.showinfo("Purge", "La purge n'a pas pu être effectuée ")
        wd.close()


## TODO SATD-jj-mm-yy.ods


# Procédure pour
def open_file():
    global File_path
    global l1
    global nb_ligne1
    source_rep = os.getcwd()
    file = filedialog.askopenfile(mode='r', filetypes=[('Ods Files', '*.ods')])
    if file:
        filepath = os.path.abspath(file.name)
        filepath = filepath.replace(os.sep, "/")
        name = os.path.basename(filepath)
        destination_rep = source_rep + '/archive_SATD/archive' + datetime.now().strftime('_%Y-%m-%d')
        if not os.path.exists(destination_rep):
            os.makedirs(destination_rep)
        label_path.configure(text="Le fichier sélectionné est : " + Path(filepath).stem)
        label_path6.configure(text="Le fichier sélectionné est : " + Path(filepath).stem)
        File_path = filepath
        shutil.copyfile(filepath, destination_rep + '/' + name)
        df = pd.read_excel(filepath)
        nb_ligne = df.shape[0]
        s = 's' if nb_ligne > 1 else ''
        messagebox.showinfo("Création d'opposition", 'Votre fichier contient ' + str(nb_ligne) + ' ligne' + s + '.')
        print('Votre fichier contient ' + str(nb_ligne) + ' ligne' + s + '.')
    filename1 = 'donnees_creation_opposition_sortie' + datetime.now().strftime('_%Y-%m-%d') + '.ods'
    filepath1 = source_rep + '/donnees_sortie/donnees_sortie' + datetime.now().strftime('_%Y-%m-%d') + '/' + filename1
    print(os.path.isfile(filepath1))
    if os.path.isfile(filepath1):
        df1 = pd.read_excel(filepath1)
        column1 = df1.columns[6]
        print("le dataframe des anciennes données : \n", df1)
        print("----------------------------------------------------------------------------")
        nb_ligne1 = df1.shape[0]
        s = 's' if nb_ligne1 > 1 else ''
        sub_df1 = df1[df1['Dossiers traités'] == 'X']
        print("le dataframe contenant les lignes déjà faites: \n", sub_df1)
        print("----------------------------------------------------------------------------")
        # if len(sub_df1) == 0:
        # messagebox.showinfo("Création d'opposition", "Aucune opération n'a été effectué pour l'instant !")
        # elif min(df.index[df1['Dossiers traités'] == 'X'].tolist()) != 0 & min(df.index[df1['Dossiers traités'] == 'X'].tolist()):
        #     premiere_partie = 'La première opposition du fichier n\'a pas été enregistré' if min(
        #         df.index[df1['Dossiers traités'] == 'X'].tolist()) == 1 else 'Les ' + str(min(
        #         df.index[df1['Dossiers traités'] == 'X'].tolist()) + 1) + 'premières oppositions du fichier n\'ont pas été ' \
        #                                                       'enregistrées '
        #     messagebox.showwarning(
        #         "Création d'opposition", 'Vous avez déjà un fichier de sortie qui contient ' + str(nb_ligne1) + ' ligne'
        #                                  + s + '.\n Une opération de création d\'opposition à déjà été lancée, '
        #                                        'l\'opération \n s\'est arrêtée à la ligne '
        #                                  + str(min(df.index[df1['Dossiers traités'] == 'X'].tolist()) + 1) + '.\n'
        #                                  + premiere_partie)
        #     print('Vous avez déjà un fichier de sortie qui contient ' + str(nb_ligne1) + ' ligne' + s +
        #           '.\n Une opération de création d\'opposition à déjà été lancée, l\'opération \n s\'est arrêtée à la '
        #           'ligne '
        #           + str(min(df.index[df1['Dossiers traités'] == 'X'].tolist()) + 1) + '.\n'
        #           + premiere_partie
        #           )
        if len(sub_df1) - len(df) != 0:
            # response = messagebox.askyesno(
            #     "Création d'opposition", "Vous avez déjà effectué les opérations sur ce fichier. Mais plusieurs lignes"
            #                              " n'ont pas été enregistré. \n")
            response = messagebox.askyesno(
                "Création d'opposition", "Le fichier a déjà été traité par l'automate, à l'exception d'une ou "
                                         "plusieurs SATD identifiée(s) par les symboles \"X, ∅, \U0001F512\" en "
                                         "colonne \"Dossiers traités\".")
            try:
                time.sleep(2)
                tab5 = Frame(tabControl, bg='#E3EBD0')
                tabControl.add(tab5, text='liste des oppositions déjà effectuées')
                df1['Date d’effet = date réception SATD'] = df['Date d’effet = date réception SATD'].dt.strftime(
                    '%d-%m-%Y')
                table = Table(tab5, dataframe=df1, read_only=True, index=FALSE)
                table.place(y=120)
                table.autoResizeColumns()
                table.show()

            except FileNotFoundError as e:
                print(e)
                messagebox.showerror('Erreur de tableau', 'Il n\'y a pas de tableau à afficher')
            if not response:
                Interface.destroy()
            else:
                pass
        else:
            response = messagebox.askyesno(
                "Création d'opposition", "Vous avez déjà effectué les opérations sur ce fichier."
                                         "\n Voulez-vous continuer")
            if not response:
                Interface.destroy()
            else:
                pass

    # else:
    #     messagebox.showinfo("Création d'opposition", "Aucune opération n'a été effectué pour l'instant !")
    file.close()
    for i in range(df.shape[0]):
        if df.loc[i].isnull().any():
            message = "la ligne {} du tableau comporte une ou plusieurs données obligatoires manquantes.\n Cette " \
                      "ligne ne sera pas traitée et sera marquée dans la colonne \"Dossiers traités\" par le symbole " \
                      "\"∅\". \n Vous pouvez renseigner les champs manquants avant de lancer l'automate.".format(i + 1)
            print(messagebox, df.loc[i])
            messagebox.showwarning("Données manquantes", message)


# Procédure pour la progress bar
def progressbar(parent):
    pb = Progressbar(parent, length=500, mode='determinate', maximum=100, value=10)
    pb.place(x=250, y=370)
    return pb


# Procédure pour la gestion de l'interface Tkinter
Interface = Tk()
Interface.geometry('1000x600')
Interface.title('SATD DGE')
paramx = 10
paramy = 170

tabControl = ttk.Notebook(Interface)
tab1 = Frame(tabControl, bg='#C7DDC5')
label1 = Label(tab1, text='Afficher la liste des oppositions', font=('Arial', 15), fg='Black', bg='#ffffff',
               relief="sunken")
label1.place(x=400, y=paramx)

creancierButton = Button(tab1, text='Afficher la liste', command=main)
creancierButton.place(x=paramx + 240, y=paramy + 40)

tab2 = Frame(tabControl, bg='#E3EBD0')
label2 = Label(tab2, text='Créer des oppositions', font=('Arial', 15), fg='Black', bg='#ffffff', relief="sunken")
label2.place(x=400, y=paramx)
# tabControl.add(tab1, text='Liste des oppositions')
# tabControl.add(tab2, text='Création des oppositions')
tabControl.pack(expand=1, fill="both")
tab6 = Frame(tabControl, bg='#E3EBD0')
tabControl.add(tab6, text='Automate SATD DGE')
tabControl.pack(expand=1, fill="both")
tab3 = Frame(tabControl, bg='#E3EBD0')
tab4 = Frame(tabControl, bg='#E3EBD0')

# Etablissement de l'image de fermeture
img = Image.open('C:/Users/meddb-jean-francoi01/Documents/Application de Creation d\'Opposition/close-button.png')
img_resize = img.resize((30, 30), Image.LANCZOS)
closeIcon = ImageTk.PhotoImage(img_resize)
closeButton1 = Button(Interface, image=closeIcon, command=lambda: tabControl.forget(tab3))
closeButton1.pack(side=LEFT)
closeButton2 = Button(Interface, image=closeIcon, command=lambda: tabControl.forget(tab4))
closeButton2.pack(side=LEFT)

EnterTable1 = StringVar()
EnterTable2 = StringVar()
EnterTable3 = StringVar()
EnterTable4 = StringVar()
EnterTable5 = StringVar()
EnterTable6 = StringVar()
EnterTable7 = StringVar()
EnterTable8 = StringVar()
EnterTable9 = StringVar()
EnterTable10 = StringVar()

labelNumeroDossier = Label(tab1, text='Numéro Dossier Opposant:', relief="sunken")
labelNumeroDossier.place(x=250, y=paramy - 30)
entryNumeroDossier = Entry(tab1, textvariable=EnterTable6, justify='center')
entryNumeroDossier.place(width=225, x=paramx + 490, y=paramy - 30)
lexique = "Lexique : \n    ● Le symbole 'X' indique que la ligne a été traitée avec succès.\n    ● Le " \
          "symbole '∅' indique que des données obligatoires sont manquantes sur la ligne en question. \n    ● Le " \
          "symbole '\U0001F512' indique que le dossier de la ligne en question et verrouillé, la ligne pourra " \
          "être retraitée dans un délai de 45 minutes.\n Pour traitées les lignes comportant des anomalies, " \
          "vous avez juste à relancé l’automates une fois les anomalies résolues. "
labelLexique = Label(tab6, text=lexique, relief="sunken", wraplength=500, justify=LEFT)
labelLexique.place(x=250, y=paramy + 235)

creerOpposition = Button(tab2, text='Créer les Oppositions avec navigateur',
                         command=lambda: create_opposition(headless=False))
creerOpposition.place(x=paramx + 240, y=paramy + 300)

label3 = Label(tab2, text='Saisir la ligne du début: ', relief="sunken")
label3.place(x=paramx + 240, y=paramy + 45)
entry2 = Entry(tab2, textvariable=EnterTable2, justify='center')
entry2.place(width=225, x=paramx + 490, y=paramy + 45)
label4 = Label(tab2, text='Saisir le nombre de lignes à traiter: ', relief="sunken")
label4.place(x=paramx + 240, y=paramy + 105)
entry3 = Entry(tab2, textvariable=EnterTable3, justify='center')
entry3.place(width=225, x=paramx + 490, y=paramy + 105)

purge_button = Button(tab2, text='Purger', command=purge)
purge_button.place(x=paramx + 240, y=paramy + 200)

browser_button = Button(tab2, text='Créer les Oppositions sans navigateur !',
                        command=lambda: create_opposition(headless=True))
browser_button.place(x=paramx + 240, y=paramy + 250)

# login et mot de passe sur tab1 à tab3
label5 = Label(tab1, text='Identifiant:', relief="sunken")
label5.place(x=250, y=70)
entry4 = Entry(tab1, textvariable=EnterTable4, justify='center')
entry4.place(x=340, y=70)
label6 = Label(tab1, text='Mot de passe: ', relief="sunken")
label6.place(x=500, y=70)
entry5 = Entry(tab1, textvariable=EnterTable5, justify='center')
entry5.place(x=600, y=70)

label5 = Label(tab2, text='Identifiant:', relief="sunken")
label5.place(x=250, y=70)
entry4 = Entry(tab2, textvariable=EnterTable4, justify='center')
entry4.place(x=340, y=70)
label6 = Label(tab2, text='Mot de passe: ', relief="sunken")
label6.place(x=500, y=70)
entry5 = Entry(tab2, textvariable=EnterTable5, justify='center')
entry5.place(x=600, y=70)

button2 = Button(tab2, text='Choisir le fichier d\'entrée', command=open_file)
button2.place(x=paramx + 240, y=paramy - 30)
label_path = Label(tab2)
label_path.place(x=paramx + 490, y=paramy - 30)

label5 = Label(tab6, text='Identifiant:', relief="sunken")
label5.place(x=250, y=70)
entry4 = Entry(tab6, textvariable=EnterTable4, justify='center')
entry4.place(x=340, y=70)
label6 = Label(tab6, text='Mot de passe: ', relief="sunken")
label6.place(x=500, y=70)
entry5 = Entry(tab6, textvariable=EnterTable5, justify='center')
entry5.place(x=600, y=70)

button2 = Button(tab6, bg="#CEDDDE", text='Choisir le fichier d\'entrée', command=open_file)
button2.place(x=paramx + 240, y=paramy - 30)
label_path6 = Label(tab6)
label_path6.place(x=paramx + 490, y=paramy - 30)

# purge_button = Button(tab6, bg="#CEDDDE", text='Purger', command=purge)
# purge_button.place(x=paramx + 240, y=paramy + 50)
# purge_label = Label(tab6, text="A utiliser en cas d'arrêt inattendu de l'automate en cours d'utilisation !",
#                     relief="sunken")
# purge_label.place(x=paramx + 340, y=paramy + 50)
browser_button = Button(tab6, bg="#82CFD8", text='Créer les Oppositions sans visualisation des transactions',
                        command=lambda: create_opposition(headless=True))
browser_button.place(x=paramx + 240, y=paramy + 100)
creerOpposition = Button(tab6, bg="#007FA9", text='Créer les Oppositions avec visualisation des transactions',
                         command=lambda: create_opposition(headless=False))
creerOpposition.place(x=paramx + 240, y=paramy + 150)

Interface.mainloop()
