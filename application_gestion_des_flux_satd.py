import os
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path
from tkinter import *
from tkinter import filedialog, messagebox, ttk, font
from tkinter.ttk import Progressbar

import numpy as np
import pandas as pd
import pyexcel_ods3 as pe
from PIL import Image, ImageTk
from pandastable import Table
from pyexcel_ods import save_data
from pynput.keyboard import Controller
from selenium import webdriver
import selenium
from selenium.common import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from _utils.error_message import ErrorMessage

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


def create_opposition(headless):
    delay = 3
    message_service_interrompu = "\nLa qualité de la connexion ne permet pas un bon fonctionnement de l'automate. " \
                                 "Veuillez essayer ultérieurement ! "

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
    filename1 = 'donnees_sortie_' + datetime.now().strftime('_%Y-%m-%d') + '.ods'
    filepath1 = source_rep + '/sorties_SATD/sorties_SATD' + datetime.now().strftime('_%Y-%m-%d') + '/' + filename1
    print("filepath1: \n", filepath1)
    print("----------------------------------------------------------------------------")
    donnees_creation_opposition = pd.read_excel(File_path)
    donnees_creation_opposition["N° dossier FRP opposé"] = donnees_creation_opposition["N° dossier FRP opposé"].astype(
        int)
    donnees_creation_opposition["N° dossier FRP opposant"] = donnees_creation_opposition[
        "N° dossier FRP opposant"].astype(int)
    donnees_creation_opposition["Montant opposition"] = donnees_creation_opposition["Montant opposition"].astype(int)

    donnees_creation_opposition["N°affaire code 1760"] = \
        donnees_creation_opposition["N°affaire code 1760"].astype(int)
    donnees_creation_opposition["Montant de l’affaire au code 1760"] = \
        donnees_creation_opposition["Montant de l’affaire au code 1760"].astype(int)
    donnees_creation_opposition["Montant à créer en « affaire service » au code 7055"] = \
        donnees_creation_opposition["Montant à créer en « affaire service » au code 7055"].astype(int)
    donnees_creation_opposition["Codique du service bénéficiaire"] = \
        donnees_creation_opposition["Codique du service bénéficiaire"].astype(str)
    donnees_creation_opposition["RANG RIB pour le remboursement du service bénéficiaire"] = \
        donnees_creation_opposition["RANG RIB pour le remboursement du service bénéficiaire"].astype(int)
    donnees_creation_opposition["RANG RIB pour le remboursement à la société "] = \
        donnees_creation_opposition["RANG RIB pour le remboursement à la société "].astype(int)
    donnees_creation_opposition["SIREN du redevable pour le libellé du virement pour la société"] = \
        donnees_creation_opposition["SIREN du redevable pour le libellé du virement pour la société"].astype(int)
    donnees_creation_opposition[
        "Numéro et date de l'opération de dépense effectuée dans Médoc pour paiement du poste comptable RNF ayant émis la SATD "] = \
        donnees_creation_opposition[
            "Numéro et date de l'opération de dépense effectuée dans Médoc pour paiement du poste comptable RNF ayant émis la SATD "].astype(
            str)
    match os.path.isfile(filepath1):
        case True:
            donnees_creation_opposition_sortie = pd.read_excel(filepath1)
            donnees_creation_opposition["Numéro d'Opération"] = ""
            donnees_creation_opposition["Date d'exécution"] = ""
            donnees_creation_opposition["Dossiers traités"] = ""
            print("dataframe des données d'entrée : \n", donnees_creation_opposition)
            print("----------------------------------------------------------------------------")
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
            print("les données après suppression des lignes incomplètes :  \n", donnees_creation_opposition)
            print("----------------------------------------------------------------------------")
            # Enlever les données déjà passées du fichier d'entrée
            old_data_done = donnees_creation_opposition_sortie[
                (donnees_creation_opposition_sortie['Dossiers traités'] == '\u2713') | (
                        donnees_creation_opposition_sortie['Dossiers traités'] == '∅') | (
                        donnees_creation_opposition_sortie['Dossiers traités'] == 'M')]
            old_data_done_list = old_data_done["Réf jugement validité = réf SATD"]
            print("liste des données déjà passées \n", old_data_done_list)
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
        case False:
            donnees_creation_opposition_sortie = pe.get_data(File_path)
            print("Mauvaise sortie")
            donnees_creation_opposition_sortie['Feuille1'][0].append("Numéro d'Opération")
            donnees_creation_opposition_sortie['Feuille1'][0].append("Date d'exécution")
            donnees_creation_opposition_sortie['Feuille1'][0].append("Dossiers traités")
            donnees_creation_opposition = pd.read_excel(File_path)

            nb_ligne = donnees_creation_opposition.shape[0]
            ligne_incomplete = list()
            satd_manuelle = list()
            last_column = "Numéro et date de l'opération de dépense effectuée dans Médoc pour paiement du poste " \
                          "comptable RNF ayant émis la SATD "
            donnees_creation_opposition['comparaison'] = donnees_creation_opposition.apply(
                lambda x: True if x[6] <= x[7] else False, axis=1)
            print("ligne 142", donnees_creation_opposition['comparaison'])
            for i in range(nb_ligne):
                # print()
                if donnees_creation_opposition.drop(columns=[last_column, 'comparaison']).loc[i].isnull().any() or \
                        donnees_creation_opposition["Date d’effet = date réception SATD"].loc[i] == 'NaT':
                    ligne_incomplete.append('∅')
                elif donnees_creation_opposition['comparaison'].loc[i]:
                    ligne_incomplete.append("M")
                    # print(ligne_incomplete)
                else:
                    ligne_incomplete.append('')
                    # print(donnees_creation_opposition.iloc[:, [7]])
            donnees_creation_opposition["Dossiers traités"] = ligne_incomplete
            print("ligne incomplete : ", ligne_incomplete)

            old_data = donnees_creation_opposition[(donnees_creation_opposition["Dossiers traités"] == '∅') | (
                    donnees_creation_opposition["Dossiers traités"] == 'M')].values \
                .tolist()
            print("les données non gardé ligne 346 \n", old_data)
            data = donnees_creation_opposition[(donnees_creation_opposition["Dossiers traités"] != '∅') & (
                    donnees_creation_opposition["Dossiers traités"] != 'M')].values.tolist()
            print("les données d'entrée ligne 347 \n", data)
            nb_ligne = len(data)
            print(nb_ligne)
    # Conversion du champ date[j][3] en string
    for j in range(nb_ligne):

        if isinstance(data[j][3], str):
            print("Ligne 559:", type(data[j][3]))
            pass
        else:
            data[j][3] = data[i][3].strftime('%d-%m-%Y')
            print("Ligne 559:", type(data[j][3]))

    print("les données d'entrée ligne 373 \n", data)
    # exit()
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
    # wd = webdriver.Edge(r"msedgedriver.exe")
    # wd = webdriver.Firefox(executable_path=GeckoDriverManager().install(), options=wd_options)
    ## TODO Passer au service object
    wd.get(
        'https://portailmetierpriv.ira.appli.impots/cas/login?service=http%3A%2F%2Fmedoc.ia.dgfip%3A8141%2Fmedocweb'
        '%2Fcas%2Fvalidation')  # adresse MEDOC DGE

    # wd.get(
    #     'http://medoc.ia.dgfip:8121/medocweb/presentation/md2oagt/ouverturesessionagent/ecran'
    #     '/ecOuvertureSessionAgent.jsf')  # adresse MEDOC Classic
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
        WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'ligneServiceHabilitation')))
    except TimeoutException:
        messagebox.showinfo("Service Interrompu !", "Le service est indisponible\n pour l'instant")
        wd.close()
    ## Saisir service
    wd.find_element(By.ID, 'nomServiceChoisi').send_keys('0070100')  # FRP MEDOC DGE
    # wd.find_element(By.ID, 'nomServiceChoisi').send_keys('6200100')
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
    ## Arriver à la transaction 3-17
    try:
        WebDriverWait(wd, 40).until(
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
        destination_rep = source_rep + '/archives_SATD/archive' + datetime.now().strftime('_%Y-%m-%d')
        num_of_secs = 60
        m, s = divmod(num_of_secs * (nb_ligne + 1), 60)
        min_sec_format = '{:02d}:{:02d}'.format(m, s)
        progressbar_label = Label(tab6,
                                  text=f"Le travail est en cours: {pb['value']:.2f}%  ~  il reste environ {min_sec_format}")
        progressbar_label.place(x=250, y=label_y)
        tab6.update()

        ## Saisie numéro de Dossier
        while True:
            print("le N° de ligne est à la ligne 309 :", j)
            print("numéro de dossier : ", data[j][0])
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))
            wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[j][0])
            wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(Keys.ENTER)
            errorMessages = ""
            print("messages d'erreur: ", errorMessages)
            time.sleep(delay)
            time.sleep(delay)
            time.sleep(delay)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
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
                j = j + 1
            else:
                break
            print("le N° de ligne est à la ligne 341 :", j)
            # exit()

        ## Saisie du choix Créer
        try:
            time.sleep(delay)
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI')))
            wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys('C')
            wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys(Keys.TAB)
            # print("ligne 473: ok")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd,delay)

        ## Saisie du numéro de dossier créancier
        try:
            time.sleep(delay)
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))
            # wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(numero_creancier_opposant)
            wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[j][1])
            wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(Keys.TAB)
            print("ligne 374: ok")
            print("le N° de ligne est à la ligne 605 :", j)  # print(data[i][1])
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)
        ## Saisie de la suite
        try:
            time.sleep(delay)
            time.sleep(delay)
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33gsuitYa33G002ReponseSuite')))
            wd.find_element(By.ID, 'inputB33gsuitYa33G002ReponseSuite').send_keys('S')
            wd.find_element(By.ID, 'inputB33gsuitYa33G002ReponseSuite').send_keys(Keys.TAB)
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        ## SAISIE DES REFERENCES DE L'OPPOSITION
        ## Transport de créance
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance').send_keys('N')
            wd.find_element(By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance').send_keys(Keys.TAB)
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        ## Saisie ATD
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GadtAdt')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GadtAdt').send_keys('O')
            wd.find_element(By.ID, 'inputB33ginf2Ya33GadtAdt').send_keys(Keys.TAB)
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        ## Saisie du crédit
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GcredCreditIs')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GcredCreditIs').send_keys('N')
            wd.find_element(By.ID, 'inputB33ginf2Ya33GcredCreditIs').send_keys(Keys.TAB)
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        ## Saisie Empêchement
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GempEmpechement')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GempEmpechement').send_keys('N')
            wd.find_element(By.ID, 'inputB33ginf2Ya33GempEmpechement').send_keys(Keys.TAB)
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        ## Saisie Montant
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GmtMontant')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GmtMontant').send_keys(data[j][2])
            wd.find_element(By.ID, 'inputB33ginf2Ya33GmtMontant').send_keys(Keys.TAB)
            # print(data[i][2])
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        ## Saisie de la Date d'Effet
        print(type(data[j][3]))
        if isinstance(data[j][3], str):
            date_d_effet = datetime.strptime(data[j][3], "%d-%m-%Y")
            print("ici c'est un string")
            print(date_d_effet.day)
        else:
            date_d_effet = data[j][3]
            print("ici ce n'est pas un string")
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour').send_keys(date_d_effet.day)
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour').send_keys(Keys.TAB)
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        ## Saisie du Mois d'Effet
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois').send_keys(date_d_effet.month)
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois').send_keys(Keys.TAB)
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        ## Saisie de l'Année d'Effet
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee').send_keys(date_d_effet.year)
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee').send_keys(Keys.TAB)
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        ## Saisie de la référence de jugement
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite').send_keys(data[j][4])
            wd.find_element(By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite').send_keys(Keys.TAB)
            # print(data[i][4])
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        ## Saisie de la date d'exécution de jugement
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementJour')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementJour').send_keys(Keys.TAB)

            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementMois')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementMois').send_keys(Keys.TAB)

            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementAnnee')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementAnnee').send_keys(Keys.TAB)
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        ## Saisie de la date de renouvellement
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementJour')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementJour').send_keys(Keys.TAB)

            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementMois')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementMois').send_keys(Keys.TAB)

            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee').send_keys(Keys.TAB)
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        ## Validation de la non saisie des dates
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee')))
            wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').send_keys('O')
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee').send_keys(Keys.TAB)
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        ## Validation de la suite
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec')))
            wd.find_element(By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec').send_keys('S')
            wd.find_element(By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec').send_keys(Keys.TAB)
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        ## Validation de la saisie de l'opposition
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33gvlcrYa33GvalcValidationCreation')))
            wd.find_element(By.ID, 'inputB33gvlcrYa33GvalcValidationCreation').send_keys('O')
            wd.find_element(By.ID, 'inputB33gvlcrYa33GvalcValidationCreation').send_keys(Keys.TAB)
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        ## Capture numéro d'opération
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'outputB33gnopeYa33GnopeNOpe')))
            numero_ope = wd.find_element(By.ID, 'outputB33gnopeYa33GnopeNOpe').text
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        ## Saisie de la fin de la phase 1bis
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33gnouvYa33GnvopNouvelleOpposition')))
            wd.find_element(By.ID, 'inputB33gnouvYa33GnvopNouvelleOpposition').send_keys('N')
            wd.find_element(By.ID, 'inputB33gnouvYa33GnvopNouvelleOpposition').send_keys(Keys.TAB)
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        ## Début de la phase 2
        time.sleep(delay)
        WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        wd.find_element(By.ID, 'barre_outils:touche_f2').click()

        ## Saisie de la transaction 21-2
        try:
            time.sleep(delay)
            WebDriverWait(wd, 10).until(
                EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi')))
            wd.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')
            time.sleep(delay)
            wd.find_element(By.ID, 'inputBmenuxBrmenx062ECaractere').send_keys('1')
            time.sleep(delay)
            WebDriverWait(wd, 10).until(
                EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi')))
            wd.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Création affaire service au code R17 "7055"
        # Saisie de la nature "AFF" pour debit 473-0
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('AFF')
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)
            print("pas 1 - ligne 722")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie du type de montant
        try:
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
            print("pas 2 - ligne 737")

        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie du montant X
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(
                data[j][7])
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(
                Keys.ENTER)
            print("pas 3 - ligne 757")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisir une identification
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located(
                    (By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep')))
            wd.find_element(By.ID,
                            'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(
                data[j][8] + " /" + data[j][4])
            wd.find_element(By.ID,
                            'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(
                Keys.ENTER)
            print("pas 4 - ligne 779")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie du numéro d'affaire
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03Nuaff1NumeroAffaire')))
            wd.find_element(By.ID, 'inputBcaff03Nuaff1NumeroAffaire').send_keys(data[j][5])
            print("pas 5 - ligne 795")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Confirmer le libelle de l'affaire
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03LaffAffLibelleAffaire')))
            wd.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(Keys.ENTER)
            print("pas 6 - ligne 811")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisir le code R27 "7370"
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi')))
            wd.find_element(By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi').send_keys(Keys.ENTER)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01RMontantSaisi')))
            time.sleep(delay)
            wd.find_element(By.ID, 'inputBcaff01Bcaff01RMontantSaisi').send_keys(Keys.ENTER)
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01SValidationOperateur')))
            wd.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys('O')
            wd.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys(Keys.ENTER)
            # WebDriverWait(wd, 40).until(
            #     EC.presence_of_element_located((By.ID, 'inputBcaff12Bcaff121ValidationON')))
            # wd.find_element(By.ID, 'inputBcaff12Bcaff121ValidationON').send_keys('O')
            # wd.find_element(By.ID, 'inputBcaff12Bcaff121ValidationON').send_keys(Keys.ENTER)
            print("pas 7 - ligne 840")
        except TimeoutException:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        # Saisir le numéro du compte 477-0
        try:
            time.sleep(delay)
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI').send_keys(Keys.ENTER)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx')))
            time.sleep(delay)
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx').send_keys('477-0')
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx').send_keys(Keys.ENTER)
            print("pas 8 - ligne 857")
        except TimeoutException:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        # Saisir la nature "AFF" pour crédit 477-0
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('AFF')
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)
            print("pas 9 - ligne 874")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie du type de montant
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
            print("pas 10 - ligne 890")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie du montant X
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(
                data[j][7])
            time.sleep(delay)
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(Keys.ENTER)
            print("pas 11 - ligne 909")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie la date comptable
        # Capture et réutilisation de la date journée comptable
        try:
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'PDATCPT_dateJourneeComptable')))
            djc_capture = wd.find_element(By.ID, 'PDATCPT_dateJourneeComptable').text
            djc = djc_capture.split('/')
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation')))
            # wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation').send_keys(djc[0])
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation').send_keys(Keys.ENTER)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation')))
            # wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation').send_keys(djc[1])
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation').send_keys(Keys.ENTER)
            # WebDriverWait(wd, 40).until(
            #     EC.presence_of_element_located((By.ID, 'affichageBandeauxDynamique:2:repeatBcimp01:0'
            #                                            ':inputBcimp01Ycimp016AnneeDateImputation')))
            # wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation').send_keys(djc[2])
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation').send_keys(Keys.ENTER)
            # time.sleep(delay)
            # wd.find_element(By.XPATH, '//*[@id="repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation"]').send_keys(
            #     Keys.ENTER)
            print("pas 12 - ligne 936")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie du numéro d'affaire
        try:
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03Nuaff1NumeroAffaire')))
            wd.find_element(By.ID, 'inputBcaff03Nuaff1NumeroAffaire').send_keys(Keys.ENTER)
            print("pas 13 - ligne 951")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie le numéro de dossier
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03Bcaff032RedevServOuRlce')))
            wd.find_element(By.ID, 'inputBcaff03Bcaff032RedevServOuRlce').send_keys('REDEV')
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03Bcaff036Car2A7NuordNumDos')))
            wd.find_element(By.ID, 'inputBcaff03Bcaff036Car2A7NuordNumDos').send_keys(data[j][0])
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03Bcaff038Cplnum')))
            wd.find_element(By.ID, 'inputBcaff03Bcaff038Cplnum').send_keys('0')
            wd.find_element(By.ID, 'inputBcaff03Bcaff038Cplnum').send_keys(Keys.ENTER)
            print("pas 14 - ligne 973")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie du libellé
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03LaffAffLibelleAffaire')))
            wd.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(
                data[j][8] + " /" + data[j][4])
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03LaffAffLibelleAffaire')))
            wd.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(Keys.ENTER)
            print("pas 15 - ligne 1000")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisir le code R27 "7055"
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi')))
            wd.find_element(By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi').send_keys(Keys.ENTER)
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Cr17R27CodeR17OuR27')))
            wd.find_element(By.ID, 'inputBcaff01Cr17R27CodeR17OuR27').send_keys('7055')
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01RMontantSaisi')))
            wd.find_element(By.ID, 'inputBcaff01Bcaff01RMontantSaisi').send_keys(data[j][7])
            wd.find_element(By.ID, 'inputBcaff01Bcaff01RMontantSaisi').send_keys(Keys.ENTER)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01SValidationOperateur')))
            wd.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys('O')
            print("pas 16 - ligne 1017")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Validation de la transaction
        try:
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcvim01Ycvim013ReponseOperateur')))
            wd.find_element(By.ID, 'inputBcvim01Ycvim013ReponseOperateur').send_keys('O')
            print("pas 17 - ligne 1032")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Création d'une liste temporaire avec numéro d'ordre de dépenses, le numéro de l'affaire créée et le numéro
        # de l'opération Le numéro de l'opération est divisé sur deux cellules dans MEDOC Cette liste sera finalement
        # collée comme ligne dans le fichier des donnees de sortie
        liste_temporaire_data = [str(data[j][0]), str(data[j][7])]  # FRP indice #0 dans liste_temporaire_data
        # #Montant indice #1 dans liste_temporaire_data

        # Numéro de l'ordre de dépense 1
        try:
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'outputBcvcs04Ycvcs022NoDepense')))
            # Numéro de l'ordre de dépense indice #2 dans liste_temporaire_data
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2')))
            liste_temporaire_data.append(wd.find_element(By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2').text)
            print("pas 18 - ligne 1056")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Pour afficher la suite
        try:
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvcs04Ycvcs028Reponse')))
            wd.find_element(By.ID, 'inputBcvcs04Ycvcs028Reponse').send_keys(Keys.ENTER)
            print("pas 19 - ligne 1070")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Numéro de l'affaire créée
        try:
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'outputBcvcs04Ycvcs021NumAffaireCreee')))
            # Numéro de l'affaire créée indice #3 dans liste_temporaire_data
            numero_affaire_creee = wd.find_element(By.ID, 'outputBcvcs04Nuaff1NumeroAffaire').text
            liste_temporaire_data.append(numero_affaire_creee)
            print("pas 20 - ligne 1087")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        if liste_temporaire_data[3] != numero_affaire_creee or liste_temporaire_data[3] == '':
            time.sleep(delay)
            numero_affaire_creee_v = wd.find_element(By.ID, 'outputBcvcs04Nuaff1NumeroAffaire').text
            liste_temporaire_data[3] = numero_affaire_creee_v
        else:
            pass

        # Pour afficher la suite
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcvcs04Ycvcs028Reponse')))
            wd.find_element(By.ID, 'inputBcvcs04Ycvcs028Reponse').send_keys(Keys.ENTER)
            print("pas 21 - ligne 1110")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Numero de l'opération 1
        try:
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'labelBcvcs032')))
            # Numero de l'opération indice #4 dans liste_temporaire_data
            liste_temporaire_data.append(
                wd.find_element(By.ID, 'outputBcvcs03Nuopet1ErCarNuopeF2').text +
                wd.find_element(By.ID, 'outputBcvcs03Nuopes5DerniersCarNuope').text)
            print("pas 22 - ligne 1127")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Pour afficher la suite
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvcs03Ycvcs014DemandeSuite')))
            wd.find_element(By.ID, 'inputBcvcs03Ycvcs014DemandeSuite').send_keys('S')
            print("pas 23 - ligne 1152")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Fin de la transaction 21-2 et retour à la page d'accueil
        WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvcs03Ycvcs014DemandeSuite')))
        wd.find_element(By.ID, 'barre_outils:image_f2').click()
        print("pas 24 - ligne 1173")

        # Saisir la transaction 21-2
        ## Saisie de la transaction 21-2
        try:
            time.sleep(delay)
            WebDriverWait(wd, 10).until(
                EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi')))
            wd.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')
            time.sleep(delay)
            wd.find_element(By.ID, 'inputBmenuxBrmenx062ECaractere').send_keys('1')
            time.sleep(delay)
            WebDriverWait(wd, 10).until(
                EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi')))
            wd.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')
            print("pas 25 - ligne 1188")
        except TimeoutException:
            progressbar_label.destroy()
            # WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            # messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messages = "Une erreur inattendu"
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        # Création affaire service au code R27 "8755"
        # Saisir la nature "AFF" pour debit 473-0
        # Saisir la nature "AFF" pour crédit 477-0
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('AFF')
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)
            print("pas 26 - ligne 1199")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie du type de montant
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
            print("pas 27 - ligne 1215")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie du montant X
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(
                data[j][7])
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(
                Keys.ENTER)
            print("pas 28 - ligne 1226")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisir une identification
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located(
                    (By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(
                data[j][8])
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(
                Keys.ENTER)
            print("pas 29 - ligne 1244")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisir le numéro d'affaire créée précédemment
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03Nuaff1NumeroAffaire')))
            wd.find_element(By.ID, 'inputBcaff03Nuaff1NumeroAffaire').send_keys(liste_temporaire_data[3])
            print("pas 30 - ligne 1260")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Confirmer le libelle de l'affaire
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03LaffAffLibelleAffaire')))
            wd.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(Keys.ENTER)
            print("pas 31 - ligne 1276")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisir le code R27 "8755"
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi')))
            wd.find_element(By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi').send_keys(Keys.ENTER)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01RMontantSaisi')))
            wd.find_element(By.ID, 'inputBcaff01Bcaff01RMontantSaisi').send_keys(Keys.ENTER)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01SValidationOperateur')))
            wd.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys('O')
            print("pas 32 - ligne 1298")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Répondre à la question "Soldez-vous l'affaire ?"
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff12Bcaff121ValidationON')))
            wd.find_element(By.ID, 'inputBcaff12Bcaff121ValidationON').send_keys('O')
            print("pas 33 - ligne 1314")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Valider CREDIT
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI').send_keys(Keys.ENTER)
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx').send_keys('512-96')
            print("pas 34 - ligne 1346")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisir la nature "OVIRT"
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('OVIRT')
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)
            print("pas 35 - ligne 1352")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisir ENTREE pour type de montant
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
            print("pas 36 - ligne 1368")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie du montant X
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(data[j][7])
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(Keys.ENTER)
            print("pas 37 - ligne 1400")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie du codique du service bénéficiaire
        try:
            time.sleep(delay)
            print("ok")
            # WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.ID, 'inputBnN4f3001Bn4F300101ZoneCodiqueService')))
            # wd.find_element(By.ID, 'inputBn4f3001Bn4F300101ZoneCodiqueService').send_keys(data[j][9])
            # script = "document.getElementById('inputBnN4f3001Bn4F300101ZoneCodiqueService').setAttribute('value', data[j][9])"
            # wd.execute_script(script)
            WebDriverWait(wd, 80).until(EC.presence_of_element_located(
                (By.XPATH, '/html/body/div[2]/div[6]/form[4]/div/div[3]/table[7]/tbody/tr[2]/td[6]/input')))
            time.sleep(delay)
            wd.find_element(By.XPATH,
                            '/html/body/div[2]/div[6]/form[4]/div/div[3]/table[7]/tbody/tr[2]/td[6]/input').send_keys(
                str(data[0][9]).rjust(2 + len(str(data[0][9])), '0'))
            print("data 9:", str(data[0][9]).rjust(2 + len(str(data[0][9])), '0'))
            print("pas 38 - ligne 1416")
        except TimeoutException:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        # Appuyer sur Entrer pour continuer
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBn4f3001Bn4F300116ZoneAcquisitionLibre')))
            wd.find_element(By.ID, 'inputBn4f3001Bn4F300116ZoneAcquisitionLibre').send_keys(Keys.ENTER)
            print("pas 39 - ligne 1435")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Validation de la transaction
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcvim01Ycvim013ReponseOperateur')))
            wd.find_element(By.ID, 'inputBcvim01Ycvim013ReponseOperateur').send_keys('O')
            print("pas 40 - ligne 1451")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Numéro de l'ordre de dépense 2
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2')))
            # Numero de l'ordre de depense indice #5 dans liste_temporaire_data
            liste_temporaire_data.append(wd.find_element(By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2').text)
            print("pas 41 - ligne 1452")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Pour afficher la suite
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcvcs04Ycvcs028Reponse')))
            wd.find_element(By.ID, 'inputBcvcs04Ycvcs028Reponse').send_keys(Keys.ENTER)
            print("pas 42 - ligne 1468")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Numéro de l'opération 2
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'labelBcvcs032')))
            # Numero de l'opération indice #6 dans liste_temporaire_data
            liste_temporaire_data.append(
                wd.find_element(By.ID, 'outputBcvcs03Nuopet1ErCarNuopeF2').text +
                wd.find_element(By.ID, 'outputBcvcs03Nuopes5DerniersCarNuope').text)

            liste_tempo_operation2_date = str(
                wd.find_element(By.ID, 'outputBcvcs03Nuopet1ErCarNuopeF2').text +
                wd.find_element(By.ID, 'outputBcvcs03Nuopes5DerniersCarNuope').text)
            print("pas 43 - ligne 1491")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Pour afficher la suite
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvcs03Ycvcs014DemandeSuite')))
            wd.find_element(By.ID, 'inputBcvcs03Ycvcs014DemandeSuite').send_keys('S')
            print("pas 44 - ligne 1506")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Fin de la transaction 21-2 et retour à la page d'accueil
        time.sleep(delay)
        WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'barre_outils:image_f2')))
        wd.find_element(By.ID, 'barre_outils:image_f2').click()
        print("pas 45 - ligne 1550")

        # Saisie la transaction 3-8-2
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi')))
            wd.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('3')
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi')))
            wd.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('8')
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi')))
            wd.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')
            print("pas 46 - ligne 1563")

        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisir le numéro d'affaire à partir des données d'entrées
        try:
            time.sleep(delay)
            time.sleep(delay)
            time.sleep(delay)
            # WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[6]/form[4]/div/div[3]/table/tbody/tr[2]/td[4]/input')))
            # wd.find_element(By.XPATH, '/html/body/div[2]/div[6]/form[4]/div/div[3]/table/tbody/tr[2]/td[4]/input').send_keys(data[j][5])
            # xpath = //*[@id="inputBrsdo03Nuaff1NumeroAffaire"]
            # xPath complet = /html/body/div[2]/div[6]/form[4]/div/div[3]/table/tbody/tr[2]/td[4]/input
            set_numero_affaire = f'''document.getElementById('inputBrsdo03Nuaff1NumeroAffaire').setAttribute('value','{data[j][5]}'); '''
            wd.execute_script(set_numero_affaire)
            time.sleep(delay)
            wd.find_element(By.XPATH,
                            '/html/body/div[2]/div[6]/form[4]/div/div[3]/table/tbody/tr[2]/td[4]/input').send_keys(
                Keys.ENTER)
            time.sleep(delay)
            time.sleep(delay)
            time.sleep(delay)
            print("pas 47 - ligne 1578")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisir le type de l'affaire "64"
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBrsdo03NasdoNatureSousDossier')))
            wd.find_element(By.ID, 'inputBrsdo03NasdoNatureSousDossier').send_keys('64')
            print("pas 48 - ligne 1593")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Récuperer le nouveau solde de l'affaire au code 1760 et enregistrer le sous indice #7 dans liste_temporaire_data
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'outputBraff01Yraff01YSoldeArticle')))
            wd.find_element(By.ID, 'outputBraff01Yraff01YSoldeArticle').text
            liste_temporaire_data.append(wd.find_element(By.ID, 'outputBraff01Yraff01YSoldeArticle').text)
            print("pas 49 - ligne 1576")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Récupérer le nom de l'entreprise à rembourser et enregistrer le dans une liste temporaire
        liste_tempo_nom_entreprise = []
        liste_tempo_nom_entreprise.append(str(data[j][0]))
        WebDriverWait(wd, 40).until(
            EC.presence_of_element_located((By.ID, 'outputBrtit04NomprfNomProfession')))
        wd.find_element(By.ID, 'outputBrtit04NomprfNomProfession').text
        liste_tempo_nom_entreprise.append(
            wd.find_element(By.ID, 'outputBrtit04NomprfNomProfession').text + "/SOLDE RCTVA")
        print("pas 50 - ligne 1594")

        # Pour afficher la suite
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'labelBrval18BarreEspace0')))
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'labelBrval18BarreEspace0')))
            wd.find_element(By.ID, 'inputYrval18wAcquisitionEspace').send_keys(Keys.ENTER)
            print("pas 51 - ligne 1602")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Pour afficher la suite encore une fois en cas de besoin
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputYrval18wAcquisitionEspace')))
            wd.find_element(By.ID, 'inputYrval18wAcquisitionEspace').send_keys(Keys.ENTER)
            print("pas 52 - ligne 1617")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Fin de la transaction 3-8-2 et retour à la page d'accueil
        time.sleep(delay)
        WebDriverWait(wd, 40).until(
            EC.presence_of_element_located((By.ID, 'barre_outils:image_f2')))
        wd.find_element(By.ID, 'barre_outils:image_f2').click()
        print("pas 53 - ligne 1631")

        # Saisir la transaction 21-2
        # Remboursement du solde à la société débitrice
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi')))
            wd.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx062ECaractere')))
            wd.find_element(By.ID, 'inputBmenuxBrmenx062ECaractere').send_keys('1')
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi')))
            wd.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')
            print("pas 54 - ligne 1645")
        except TimeoutException:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()

        # Saisir la nature "AFF" pour debit 473-0
        try:
            WebDriverWait(wd, 10).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('AFF')
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)
            print("pas 55 - ligne 1661")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie du type de montant
        try:
            WebDriverWait(wd, 10).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
            print("pas 56 - ligne 1676")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie du montant X
        montant = liste_temporaire_data[7].replace('+', '')
        # WebDriverWait(wd, 10).until(
        #     EC.presence_of_element_located(
        #         (By.XPATH, '/html/body/div[2]/div[6]/form[4]/div/div[3]/table[1]/tbody/tr[3]/td[18]/input')))
        # wd.find_element(By.XPATH,
        #                 '/html/body/div[2]/div[6]/form[4]/div/div[3]/table[1]/tbody/tr[3]/td[18]/input').send_keys(
        #     montant)
        time.sleep(delay)
        time.sleep(delay)
        time.sleep(delay)
        print(montant)
        set_montant = f'''document.getElementById('repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').setAttribute('value','{montant}');'''
        wd.execute_script(set_montant)
        print(montant)
        time.sleep(delay)
        wd.find_element(By.XPATH,
                        '/html/body/div[2]/div[6]/form[4]/div/div[3]/table[1]/tbody/tr[3]/td[18]/input').send_keys(
            Keys.ENTER)

        print("pas 57 - ligne 1693")
        # try:
        #     montant = liste_temporaire_data[7].replace('+', '')
        #     # set_montant = 'document.getElementById("repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation' \
        #     #               '").setAttribute("value","{montant}") '
        #     # wd.execute_script(set_montant)
        #     # WebDriverWait(wd, 80).until(
        #     #     EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation')))
        #     # wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(
        #     #     montant)
        #     WebDriverWait(wd, 80).until(
        #         EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[6]/form[4]/div/div[3]/table[1]/tbody/tr[3]/td[18]/input')))
        #     wd.find_element(By.XPATH, '/html/body/div[2]/div[6]/form[4]/div/div[3]/table[1]/tbody/tr[3]/td[18]/input').send_keys(
        #         montant)
        #     print(montant)
        #     wd.find_element(By.XPATH, '/html/body/div[2]/div[6]/form[4]/div/div[3]/table[1]/tbody/tr[3]/td[18]/input').send_keys(Keys.ENTER)
        #     print("pas 57 - ligne 1693")
        # except:
        #     progressbar_label.destroy()
        #     WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
        #     messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
        #     messagebox.showinfo("Service Interrompu !", messages)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.close()

        # Saisie de l'identification
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located(
                    (By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(
                liste_tempo_nom_entreprise[1])
            time.sleep(delay)
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(
                Keys.ENTER)
            print("pas 58 - ligne 1710")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie du numéro d'affaire créée précédemment
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03Nuaff1NumeroAffaire')))
            wd.find_element(By.ID, 'inputBcaff03Nuaff1NumeroAffaire').send_keys(data[j][5])
            print("pas 59 - ligne 1725")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Confirmer le libelle de l'affaire
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03LaffAffLibelleAffaire')))
            wd.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(Keys.ENTER)
            print("pas 60 - ligne 1740")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisir le code R27 "7370"
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi')))
            wd.find_element(By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi').send_keys(Keys.ENTER)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01RMontantSaisi')))
            wd.find_element(By.ID, 'inputBcaff01Bcaff01RMontantSaisi').send_keys(Keys.ENTER)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01SValidationOperateur')))
            wd.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys('O')
            print("pas 61 - ligne 1761")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Répondre à la question "Soldez-vous l'affaire ?"
        try:
            time.sleep(delay)
            WebDriverWait(wd, 60).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff12Bcaff121ValidationON')))
            wd.find_element(By.ID, 'inputBcaff12Bcaff121ValidationON').send_keys('O')
            print("pas 62 - ligne 1777")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Valider CREDIT
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI').send_keys(Keys.ENTER)
            print("pas 63 - ligne 1911")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisir le numéro du compte 512-96
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx').send_keys('512-96')
            print("pas 64 - ligne 1807")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie de la nature "VIRT"
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('VIRT')
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)
            print("pas 65 - ligne 1823")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie du type de montant
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
            print("pas 66 - ligne 1838")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie du montant X
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(
                liste_temporaire_data[7])
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(Keys.ENTER)
            print("pas 67 - ligne 1855")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie de la date du jour comptable
        try:
            time.sleep(delay)
            # WebDriverWait(wd, 40).until(
            #     EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation')))
            # wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation').send_keys(djc[0])
            # WebDriverWait(wd, 40).until(
            #     EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation')))
            # wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation').send_keys(djc[1])
            # WebDriverWait(wd, 40).until(
            #     EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation')))
            # wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation').send_keys(djc[2])
            # wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation').send_keys(Keys.ENTER)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation').send_keys(Keys.ENTER)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation').send_keys(Keys.ENTER)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation').send_keys(Keys.ENTER)
            print("pas 68 - ligne 1876")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Saisie du numéro de dossier
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))
            wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[j][0])
            print("pas 69 - ligne 1891")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Continuer en cas d'existence de ce message : ATTENTION - OPPOSITION POUR CE DOSSIER
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_all_elements_located((By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON')))
            wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').send_keys('O')
            print("pas 70 - ligne 1906")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Continuer en cas d'existence de RAR
        try:
            if wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').is_displayed:
                time.sleep(delay)
                wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').send_keys('O')
        except:
            pass

        # Saisie du numéro de l'IBAN
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_all_elements_located((By.ID, 'inputBibanremYaribmess1LibelleMessage')))
            wd.find_element(By.ID, 'inputBibanremYaribchoixSaisieChoix').send_keys(data[j][11])
            print("pas 71 - ligne 1929")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Libelle du virement emis
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_all_elements_located((By.ID, 'repeatBcimp01:1:inputBcrib01IbanlibLibelleVirementEmis')))
            wd.find_element(By.ID, 'repeatBcimp01:1:inputBcrib01IbanlibLibelleVirementEmis').send_keys(
                str(data[j][12]) + "/ RCTVA")
            wd.find_element(By.ID, 'repeatBcimp01:1:inputBcrib01IbanlibLibelleVirementEmis').send_keys(Keys.ENTER)
            print("pas 72 - ligne 1946")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Répondre à la question "Voulez-vous valider ?"
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcvim01Ycvim013ReponseOperateur')))
            wd.find_element(By.ID, 'inputBcvim01Ycvim013ReponseOperateur').send_keys('O')
            print("pas 73 - ligne 1961")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Visualisation du Numero de l'ordre de dépense
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2')))
            # Numero de l'ordre de depense indice #8 dans liste_temporaire_data
            liste_temporaire_data.append(wd.find_element(By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2').text)
            print("pas 73 - ligne 1977")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Pour afficher la suite
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcvcs04Ycvcs028Reponse')))
            wd.find_element(By.ID, 'inputBcvcs04Ycvcs028Reponse').send_keys(Keys.ENTER)
            print("pas 74 - ligne 1992")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Numero de l'opération
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'labelBcvcs032')))
            # Numero de l'opération indice #9 dans liste_temporaire_data
            liste_temporaire_data.append(
                wd.find_element(By.ID, 'outputBcvcs03Nuopet1ErCarNuopeF2').text +
                wd.find_element(By.ID, 'outputBcvcs03Nuopes5DerniersCarNuope').text)
            print("pas 75 - ligne 2009")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Pour afficher la suite
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvcs03Ycvcs014DemandeSuite')))
            wd.find_element(By.ID, 'inputBcvcs03Ycvcs014DemandeSuite').send_keys('S')
            print("pas 76 - ligne 2023")
        except TimeoutException:
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.close()
        except:
            progressbar_label.destroy()
            ErrorMessage.error_message(wd, delay)

        # Fin de la transaction 21-2 et retour à la page d'accueil
        WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'barre_outils:image_f2')))
        wd.find_element(By.ID, 'barre_outils:image_f2').click()

        # convertir les colonnes numériques de la liste data en entier

        # exit()
        ## Marquage tâche faîte dans le fichier
        match os.path.isfile(filepath1):
            case True:
                data[j][14] = numero_ope
                data[j][15] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                data[j][16] = '\u2713'
                print("inscription des données dans la liste ligne 929", data)
            case False:
                data[j][3] = str(date_d_effet.strftime('%Y-%m-%d'))
                data[j].insert(14, numero_ope)
                data[j].insert(15, datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
                data[j][16] = '\u2713'
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
        if j + 1 < nb_ligne:
            j += 1
        else:
            break
        columns = ["FRP société", "FRP opposant", "Montant opposition","Date d’effet = date réception SATD",
                   "Réf jugement validité = réf SATD", "N°affaire code 1760",
                   "Montant de l’affaire au code 1760", "Montant à créer en «affaire service» au code 7055",
                   "Identification du bénéficiaire de la dépense", "Codique du service bénéficiaire",
                   "RANG RIB pour le remboursement du service bénéficiaire",
                   "RANG RIB pour le remboursement à la société",
                   "SIREN du redevable pour le libellé du virement pour la société",
                   "Numéro et date de l'opération de dépense effectuée dans Médoc pour paiement du poste "
                   "comptable RNF ayant émis la SATD ", "Numéro d'Opération","Date d'exécution", "Dossiers traités"]
        data.insert(0, columns)
        print("les nouvelles data : \n", data)
        # source_rep = os.getcwd()
        destination_rep1 = source_rep + '/sorties_SATD/sorties_SATD' + datetime.now().strftime('_%Y-%m-%d')
        saved_file = destination_rep1 + '/' + filename1
        if not os.path.exists(destination_rep1):
            os.makedirs(destination_rep1)
        if os.path.exists(destination_rep1 + '/' + filename1):
            os.remove(destination_rep1 + '/' + filename1)
            for i in range(len(old_data)):
                old_data[i].insert(14, '')
                old_data[i].insert(15, '')
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
                if old_data[i][14] == '' & old_data[i][15] == '':
                    del old_data[i][14]
                    del old_data[i][15]
            print("old_data : \n", old_data)
            del data[0]
            print("data sans les entêtes (ligne 993)", data)
            if not old_data:
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


# Procédure pour l'ouverture des fichiers d'entrées
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
        destination_rep = source_rep + '/archives_SATD/archive' + datetime.now().strftime('_%Y-%m-%d')
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
    filename1 = 'donnees_sortie_' + datetime.now().strftime('_%Y-%m-%d') + '.ods'
    filepath1 = source_rep + '/sorties_SATD/sorties_SATD' + datetime.now().strftime('_%Y-%m-%d') + '/' + filename1
    print(os.path.isfile(filepath1))
    if os.path.isfile(filepath1):
        df1 = pd.read_excel(filepath1)
        df1 = df1.fillna(0)
        column1 = df1.columns[6]
        print("le dataframe des anciennes données : \n", df1)
        print("----------------------------------------------------------------------------")
        nb_ligne1 = df1.shape[0]
        s = 's' if nb_ligne1 > 1 else ''
        sub_df1 = df1[df1['Dossiers traités'] == 'X']
        print("le dataframe contenant les lignes déjà faites: \n", sub_df1)
        print("----------------------------------------------------------------------------")
        if len(sub_df1) - len(df) != 0:
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
                "Automate SATD", "Vous avez déjà effectué les opérations sur ce fichier."
                                 "\n Voulez-vous continuer")
            if not response:
                Interface.destroy()
            else:
                pass

    # else:
    #     messagebox.showinfo("Création d'opposition", "Aucune opération n'a été effectué pour l'instant !")
    print(df)
    for i in range(df.shape[0]):
        last_colonne = "Numéro et date de l'opération de dépense effectuée dans Médoc pour paiement du poste comptable RNF ayant émis la SATD "
        if df.drop(columns=[last_colonne]).loc[i].isnull().any():
            message = "la ligne {} du tableau comporte une ou plusieurs données obligatoires manquantes.\n Cette " \
                      "ligne ne sera pas traitée et sera marquée dans la colonne \"Dossiers traités\" par le symbole " \
                      "\"∅\". \n Vous pouvez renseigner les champs manquants avant de lancer l'automate.".format(i + 1)
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

tab2 = Frame(tabControl, bg='#E3EBD0')
label2 = Label(tab2, text='Créer des oppositions', font=('Arial', 15), fg='Black', bg='#ffffff', relief="sunken")
label2.place(x=400, y=paramx)
# tabControl.add(tab1, text='Liste des oppositions')
# tabControl.add(tab2, text='Création des oppositions')
tabControl.pack(expand=1, fill="both")
tab6 = Frame(tabControl, bg='#F1F1D3')
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
buttonFont = font.Font(family='Tahoma', size=15)
question = '\U0000003F'

lexique = "Précisions sur les symboles affichés en colonne \"Dossiers traités\" du fichier de sortie :\n" \
          "\n● Le symbole \"\u2713\" indique que la SATD a été traitée jusqu'à la mainlevée." \
          "\n● Le symbole \"∅\" indique qu'une ou plusieurs données obligatoires sont manquantes sur la ligne, ce qui " \
          "ne permet pas de traiter la SATD. Il convient de compléter la ou les données manquantes avant d'exécuter de" \
          " nouveau l'automate pour traiter la SATD concernée." \
          "\n● Le symbole \"\U0001F512\" indique que le dossier FRP de l'opposé est verrouillé dans MEDOC. Il convient" \
          " d'attendre un délai de 45 minutes avant d'exécuter de nouveau l'automate pour traiter la SATD concernée. " \
          "\n● Le symbole \"M\" indique que l'automate n'est pas en mesure de traiter la SATD. Le traitement doit " \
          "être effectué manuellement. "

lexiqueButton = Button(tab6, bg="#E3EBD0", text=question, font=buttonFont,
                       command=lambda: messagebox.showinfo("Indicateurs du fichier de sortie", lexique))
lexiqueButton.place(x=250, y=paramy + 50)
labelNumeroDossier = Label(tab1, text='Numéro Dossier Opposant:', relief="sunken")
labelNumeroDossier.place(x=250, y=paramy - 30)
entryNumeroDossier = Entry(tab1, textvariable=EnterTable6, justify='center')
entryNumeroDossier.place(width=225, x=paramx + 490, y=paramy - 30)

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
browser_button = Button(tab6, bg="#C7DDC5", text='Créer les Oppositions sans visualisation des transactions',
                        command=lambda: create_opposition(headless=True))
browser_button.place(x=paramx + 240, y=paramy + 100)
creerOpposition = Button(tab6, bg="#9FCDA8", text='Créer les Oppositions avec visualisation des transactions',
                         command=lambda: create_opposition(headless=False))
creerOpposition.place(x=paramx + 240, y=paramy + 150)

Interface.mainloop()
