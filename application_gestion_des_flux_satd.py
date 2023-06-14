import logging
import os
import shutil
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from sys import exit
from tkinter import *
from tkinter import filedialog, messagebox, ttk, font
from tkinter.ttk import Progressbar

import numpy as np
import pandas as pd
from pandastable import Table
from pynput.mouse import Controller
from pyexcel_ods import save_data
from selenium import webdriver
from selenium.common import TimeoutException, StaleElementReferenceException, ElementNotInteractableException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from _utils.error_message import ErrorMessage
from _utils.save_file import Saved_file

sys.stdin.reconfigure(encoding='utf-8')
sys.stdout.reconfigure(encoding='utf-8')
# import logging

mouse = Controller()
global success
success = '✓'
global vide
vide = "∅"

global enter
enter = "new KeyboardEvent('keydown', {altKey:false,bubbles: true, cancelBubble: false,cancelable: true," \
        "charCode: 0,code: 'Enter',composed: true,ctrlKey: false,currentTarget: null, defaultPrevented: true, " \
        "detail: 0, eventPhase: 0, isComposing: false, isTrusted: true, key: 'Enter', keyCode: 13, location: 0, " \
        "metaKey: false, repeat: false, returnValue: false, shiftKey: false, type: 'keydown', which: 13}); "


def __init__(self, progress):
    self.progress = progress


# Fonction pour retrouver le chemin d'accès
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


def create_opposition(headless):
    logging.basicConfig(filename=f'''automate_SATD{datetime.now().strftime('_%Y-%m-%d-%H_%M%S')}.log ''', filemode='w',
                        format='%(name)s - %(levelname)s - %(message)s', datefmt='%d-%m-%y %H:%M:%S',
                        level=logging.INFO)
    heure_demarrage = datetime.now()
    print("heure de démarrage de l'application : ", heure_demarrage.time().strftime('%H:%M:%S'))
    logging.info(f'''heure de démarrage de l'application : "{heure_demarrage.time().strftime('%H:%M:%S')}");''')
    delay = 4
    err = ErrorMessage()
    sav = Saved_file()
    message_service_interrompu = "\nLa qualité de la connexion ne permet pas un bon fonctionnement de l'automate. " \
                                 "Veuillez essayer ultérieurement ! "
    # Definition du repertoire de sauvegarde du fichier de sortie
    source_rep = os.getcwd()
    fichier_de_Sortie = 'donnees_sortie' + datetime.now().strftime('_%Y-%m-%d') + '.ods'
    sortie_repertoire = source_rep + '/sorties_SATD/sorties_SATD' + datetime.now().strftime('_%Y-%m-%d')
    saved_file = sortie_repertoire + '/' + fichier_de_Sortie

    # Configuration de la barre de progression
    pb = progressbar(tab6)
    progressbar_label = Label(tab6, text=f"Le travail commence. L'automate se connecte...")
    label_y = 390
    progressbar_label.place(x=250, y=label_y)
    tab6.update()

    time.sleep(delay)

    ## Saisie du nom utilisateur et mot de passe

    while True:
        login = EnterTable4.get()
        if login:  ##vérifie que ça soit un numéro
            break
        else:
            messagebox.showerror("Erreur de saisie", 'Saisie incorrecte, réessayez')
            exit()

    # Combien de lignes du fichier traiter
    while True:
        mot_de_passe = EnterTable5.get()
        if mot_de_passe:
            break
        else:
            messagebox.showerror("Erreur de saisie", 'Saisie incorrecte, réessayez')
            exit()

    ## Prend les données depuis le fichier, crée une liste de listes (ou "array"), oú chaque liste est
    ## une ligne du fichier Calc. Il faut faire ça parce que pyxcel_ods prend les données sous forme
    ## de dictionnaire.

    filepath1 = source_rep + '/sorties_SATD/sorties_SATD' + datetime.now().strftime(
        '_%Y-%m-%d') + '/' + fichier_de_Sortie
    print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Automate SATD ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    logging.info(
        "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ Automate SATD ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

    print("----------------------------------------------------------------------------------------------------------")
    logging.info(
        "----------------------------------------------------------------------------------------------------------")
    donnees_creation_opposition = pd.read_excel(File_path).fillna(0)
    columns = donnees_creation_opposition.columns.values.tolist()
    del columns[14]
    columns_sortie = columns + ["Numéro d'Opération", "Date d'exécution", "Dossiers traités"]

    donnees_creation_opposition["N° dossier FRP opposé"] = donnees_creation_opposition["N° dossier FRP opposé"].astype(
        str, errors='ignore')
    donnees_creation_opposition["N° dossier FRP opposant"] = donnees_creation_opposition[
        "N° dossier FRP opposant"].astype(
        str, errors='ignore')
    donnees_creation_opposition["Montant opposition (en euros)"] = donnees_creation_opposition[
        "Montant opposition (en euros)"].astype(
        int, errors='ignore')
    donnees_creation_opposition["N°affaire code 1760"] = \
        donnees_creation_opposition["N°affaire code 1760"].astype(
            int, errors='ignore')
    donnees_creation_opposition["Montant de l’affaire au code 1760 (en euros)"] = \
        donnees_creation_opposition["Montant de l’affaire au code 1760 (en euros)"].astype(
            int, errors='ignore')
    donnees_creation_opposition["Montant à créer en « affaire service » au code 7055 (en euros)"] = \
        donnees_creation_opposition["Montant à créer en « affaire service » au code 7055 (en euros)"].astype(
            int, errors='ignore')
    donnees_creation_opposition["Codique du service bénéficiaire"] = \
        donnees_creation_opposition["Codique du service bénéficiaire"].astype(str)
    donnees_creation_opposition["RANG RIB pour le remboursement du service bénéficiaire"] = \
        donnees_creation_opposition["RANG RIB pour le remboursement du service bénéficiaire"].astype(
            int, errors='ignore')
    donnees_creation_opposition["RANG RIB pour le remboursement à la société "] = \
        donnees_creation_opposition["RANG RIB pour le remboursement à la société "].astype(
            int, errors='ignore')
    donnees_creation_opposition["SIREN du redevable pour le libellé du virement pour la société"] = \
        donnees_creation_opposition["SIREN du redevable pour le libellé du virement pour la société"].astype(
            int, errors='ignore')
    donnees_creation_opposition[
        "Numéro et date de l'opération de dépense effectuée dans Médoc pour paiement du poste comptable RNF ayant émis la SATD "] = \
        donnees_creation_opposition[
            "Numéro et date de l'opération de dépense effectuée dans Médoc pour paiement du poste comptable RNF ayant émis la SATD "].astype(
            str)
    last_column = "Numéro et date de l'opération de dépense effectuée dans Médoc pour paiement du poste " \
                  "comptable RNF ayant émis la SATD "
    match os.path.isfile(filepath1):
        case True:
            donnees_creation_opposition_sortie = pd.read_excel(filepath1)
            donnees_creation_opposition_sortie["N° dossier FRP opposé"] = donnees_creation_opposition_sortie[
                "N° dossier FRP opposé"].astype(
                str, errors='ignore')
            donnees_creation_opposition_sortie["N° dossier FRP opposant"] = donnees_creation_opposition_sortie[
                "N° dossier FRP opposant"].astype(
                str, errors='ignore')
            donnees_creation_opposition_sortie["Montant opposition (en euros)"] = donnees_creation_opposition_sortie[
                "Montant opposition (en euros)"].astype(
                int, errors='ignore')
            donnees_creation_opposition_sortie["Numéro de facture Chorus"] = donnees_creation_opposition_sortie[
                "Numéro de facture Chorus"].astype(
                str, errors='ignore')
            donnees_creation_opposition_sortie["N°affaire code 1760"] = \
                donnees_creation_opposition_sortie["N°affaire code 1760"].astype(
                    int, errors='ignore')
            donnees_creation_opposition_sortie["Montant de l’affaire au code 1760 (en euros)"] = \
                donnees_creation_opposition_sortie["Montant de l’affaire au code 1760 (en euros)"].astype(
                    int, errors='ignore')
            donnees_creation_opposition_sortie["Montant à créer en « affaire service » au code 7055 (en euros)"] = \
                donnees_creation_opposition_sortie[
                    "Montant à créer en « affaire service » au code 7055 (en euros)"].astype(
                    int, errors='ignore')
            donnees_creation_opposition_sortie["Identification du bénéficiaire de la dépense"] = \
                donnees_creation_opposition_sortie["Identification du bénéficiaire de la dépense"].astype(str)
            donnees_creation_opposition_sortie["Codique du service bénéficiaire"] = \
                donnees_creation_opposition_sortie["Codique du service bénéficiaire"].astype(str)
            donnees_creation_opposition_sortie["RANG RIB pour le remboursement du service bénéficiaire"] = \
                donnees_creation_opposition_sortie["RANG RIB pour le remboursement du service bénéficiaire"].astype(
                    int, errors='ignore')
            donnees_creation_opposition_sortie["RANG RIB pour le remboursement à la société "] = \
                donnees_creation_opposition_sortie["RANG RIB pour le remboursement à la société "].astype(
                    int, errors='ignore')
            donnees_creation_opposition_sortie["SIREN du redevable pour le libellé du virement pour la société"] = \
                donnees_creation_opposition_sortie[
                    "SIREN du redevable pour le libellé du virement pour la société"].astype(
                    int, errors='ignore')
            donnees_creation_opposition_sortie[
                "Numéro et date de l'opération de dépense effectuée dans Médoc pour paiement du poste comptable RNF ayant émis la SATD "] = \
                donnees_creation_opposition_sortie[
                    "Numéro et date de l'opération de dépense effectuée dans Médoc pour paiement du poste comptable RNF ayant émis la SATD "].astype(
                    str)
            # donnees_creation_opposition_sortie["Numéro d'Opération"] = donnees_creation_opposition_sortie[
            #     "Numéro d'Opération"].astype(str)
            donnees_creation_opposition_sortie["Date d'exécution"] = donnees_creation_opposition_sortie[
                "Date d'exécution"].astype(str)
            # donnees_creation_opposition_sortie["Dossiers traités"] = donnees_creation_opposition_sortie[
            #     "Dossiers traités"].astype(str)
            print("Dossier sortie existant")
            donnees_creation_opposition = pd.read_excel(File_path)
            # donnees_creation_opposition.insert(14, "Numéro d'Opération", "0", allow_duplicates=False)
            # donnees_creation_opposition["Numéro d'Opération"] = '0'
            # donnees_creation_opposition["Numéro d'Opération"] = donnees_creation_opposition[
            #     "Numéro d'Opération"].astype(str)
            # donnees_creation_opposition["Date d'exécution"] = '0'
            donnees_creation_opposition.insert(15, "Date d'exécution", "0", allow_duplicates=False)
            donnees_creation_opposition["Date d'exécution"] = donnees_creation_opposition[
                "Date d'exécution"].astype(str)
            donnees_creation_opposition["Dossiers traités"] = '0'
            donnees_creation_opposition["Dossiers traités"] = donnees_creation_opposition[
                "Dossiers traités"].astype(str)
            nb_ligne = donnees_creation_opposition.shape[0]
            ligne_incomplete = list()
            print("---------------------------------------------------------------------------------------------------")
            taille_donnee_entree = donnees_creation_opposition.shape[0]
            donnees_creation_opposition['comparaison'] = donnees_creation_opposition.apply(
                lambda x: True if x[6] <= x[7] else False, axis=1)
            print("ligne 142", donnees_creation_opposition['comparaison'])
            for i in range(nb_ligne):
                # print()
                if donnees_creation_opposition.drop(columns=[last_column, 'comparaison']).loc[i].isnull().any() or \
                        donnees_creation_opposition["Date d’effet = date réception SATD"].loc[i] == 'NaT':
                    ligne_incomplete.append(vide)
                elif donnees_creation_opposition['comparaison'].loc[i] or \
                        donnees_creation_opposition["Numéro de facture Chorus"].duplicated().loc[i]:
                    ligne_incomplete.append("M")
                    # print(ligne_incomplete)
                else:
                    ligne_incomplete.append('0')
                    # print(donnees_creation_opposition.iloc[:, [7]])
            donnees_creation_opposition["Dossiers traités"] = ligne_incomplete
            print("ligne incomplete : ", ligne_incomplete)
            donnees_creation_opposition.drop(["comparaison"], axis=1, inplace=True)
            bad_df = donnees_creation_opposition[(donnees_creation_opposition["Dossiers traités"] == vide) | (
                    donnees_creation_opposition["Dossiers traités"] == 'M')].fillna(0)
            old_data_df = pd.concat([donnees_creation_opposition_sortie, bad_df])
            old_data_df.drop_duplicates(subset=None, keep='first', inplace=False)
            old_data_df["Date d’effet = date réception SATD"] = pd.to_datetime(
                old_data_df["Date d’effet = date réception SATD"], format='%Y-%m-%d')
            old_data_df_done = old_data_df[
                (old_data_df["Dossiers traités"] == success) | (old_data_df["Dossiers traités"] == 'M') | (
                        old_data_df["Dossiers traités"] != '0')]
            print("old_data_df_done", old_data_df_done["N°affaire code 1760"])
            data_df = donnees_creation_opposition[(donnees_creation_opposition["Dossiers traités"] != vide) & (
                    donnees_creation_opposition["Dossiers traités"] != 'M')].fillna(0)
            data_df["filter"] = data_df["N°affaire code 1760"].isin(old_data_df_done["N°affaire code 1760"])

            print("data_df", data_df["filter"])
            data_df = data_df[(data_df["filter"] == False)]
            old_data_df["filter2"] = old_data_df["N°affaire code 1760"].isin(data_df["N°affaire code 1760"])
            old_data_df = old_data_df[(old_data_df["filter2"] == False)]
            old_data_df.drop(["filter2"], axis=1, inplace=True)
            data_df.drop(["filter"], axis=1, inplace=True)
            old_data = old_data_df.values.tolist()
            print("les données non gardé ligne 236 \n", old_data)
            data = data_df.values.tolist()
            for i in range(len(data)):
                data[i][0] = str(data[i][0]).replace(" ", "")
                data[i][1] = str(data[i][1]).replace(" ", "")
                data[i][9] = str(data[i][9]).replace(" ", "")
                data[i][12] = str(data[i][12]).replace(" ", "")
            print("les données d'entrée ligne 189 \n", data)
            if not data:
                messagebox.showinfo("Données Manquante", "Il n'y a pas de données exploitables. "
                                                         "Veuillez vérifier le fichiers d'entrée")
                exit()
            else:
                pass
            nb_ligne = len(data)
            print(nb_ligne)
        case False:
            print("pas de dossier sortie existant")
            donnees_creation_opposition = pd.read_excel(File_path)
            donnees_creation_opposition["Dossiers traités"] = '0'
            # donnees_creation_opposition.insert(14, "Numéro d'Opération", "0", allow_duplicates=False)
            # donnees_creation_opposition["Numéro d'Opération"] = donnees_creation_opposition[
            #     "Numéro d'Opération"].astype(str)
            donnees_creation_opposition.insert(15, "Date d'exécution", "0", allow_duplicates=False)
            donnees_creation_opposition["Date d'exécution"] = donnees_creation_opposition[
                "Date d'exécution"].astype(str)
            nb_ligne = donnees_creation_opposition.shape[0]
            ligne_incomplete = list()
            donnees_creation_opposition['comparaison'] = donnees_creation_opposition.apply(
                lambda x: True if x[6] <= x[7] else False, axis=1)
            for i in range(nb_ligne):
                # print()
                if donnees_creation_opposition.drop(columns=[last_column, 'comparaison']).loc[i].isnull().any() or \
                        donnees_creation_opposition["Date d’effet = date réception SATD"].loc[i] == 'NaT':
                    ligne_incomplete.append(vide)
                elif donnees_creation_opposition['comparaison'].loc[i] or \
                        donnees_creation_opposition["Numéro de facture Chorus"].duplicated().loc[i]:
                    ligne_incomplete.append("M")
                    # print(ligne_incomplete)
                else:
                    ligne_incomplete.append('0')
                    # print(donnees_creation_opposition.iloc[:, [7]])
            donnees_creation_opposition["Dossiers traités"] = ligne_incomplete
            # print("ligne incomplete :\n", ligne_incomplete)
            donnees_creation_opposition.drop(["comparaison"], axis=1, inplace=True)
            # print("data ligne 302:\n", donnees_creation_opposition)
            old_df = donnees_creation_opposition[(donnees_creation_opposition["Dossiers traités"] == vide) | (
                    donnees_creation_opposition["Dossiers traités"] == 'M')]
            old_data = old_df.values.tolist()
            print("les données non gardé ligne 306 \n", old_data)
            data = donnees_creation_opposition[(donnees_creation_opposition["Dossiers traités"] != vide) & (
                    donnees_creation_opposition["Dossiers traités"] != 'M')].values.tolist()
            for i in range(len(data)):
                data[i][0] = str(data[i][0]).replace(" ", "")
                data[i][1] = str(data[i][1]).replace(" ", "")
                data[i][9] = str(data[i][9]).replace(" ", "")
                data[i][12] = str(data[i][12]).replace(" ", "")
            print("les données d'entrée ligne 310 \n", data)
            if not data:
                messagebox.showinfo("Données Manquante", "Il n'y a pas de données exploitables. "
                                                         "Veuillez vérifier le fichiers d'entrée")
                exit()
            else:
                pass
            nb_ligne = len(data)
            print(nb_ligne)
    # Conversion du champ date[j][3] en string
    for j in range(nb_ligne):
        if isinstance(data[j][3], str):
            print("Ligne 322:", type(data[j][3]))
            pass
        else:
            data[j][3] = data[j][3].strftime('%Y-%m-%d')
            print("Ligne 326:", type(data[j][3]))
    if old_data:
        for k in range(len(old_data)):
            if isinstance(old_data[k][3], str) or old_data[k][3] == 0:
                print("Ligne 262 :", type(old_data[k][3]))
                pass
            else:
                old_data[k][3] = old_data[k][3].strftime('%Y-%m-%d')
                print("Ligne 334:", type(old_data[k][3]))
    # exit()
    wd_options = webdriver.FirefoxOptions()
    if headless:
        wd_options.add_argument("--headless")
    else:
        pass
    wd_options.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'

    wd_options.set_preference('detach', True)
    wd = webdriver.Firefox(options=wd_options)
    # url = wd.command_executor._url
    # session_id = wd.session_id
    # print("l'url courant est :", url)
    # print("la session est :", session_id)
    # exit()
    wd.get(
        'https://portailmetierpriv.ira.appli.impots/cas/login?service=http%3A%2F%2Fmedoc.ia.dgfip%3A8141%2Fmedocweb'
        '%2Fcas%2Fvalidation')  # adresse MEDOC DGE

    # wd.get('https://portailmetierpriv.appli.impots/oauth2/authorize?response_type=code&redirect_uri=https%3A%2F%2Fauth-portail.appli.impots%2F%3Fopenidconnectcallback%3D1&nonce=1682343970_19914&client_id=dgfip&display=&state=1682343970_17561&scope=openid%20profile')  # adresse MEDOC DGE Réel

    # wd.get(
    #     'http://medoc.ia.dgfip:8121/medocweb/presentation/md2oagt/ouverturesessionagent/ecran'
    #     '/ecOuvertureSessionAgent.jsf')  # adresse MEDOC Classic

    print(wd.title)
    while wd.title == "Identification":
        ##Saisir utilisateur
        time.sleep(delay)
        # script = f'''identifant = document.getElementById('identifiant'); identifiant.setAttribute('type','hidden'); identifiant.setAttribute('value',"{login}");'''
        # script = f'''identifant = document.getElementById('identifiant'); identifiant.setAttribute('value',"{login}");'''
        time.sleep(delay)
        wd.find_element(By.ID, 'identifiant').send_keys(login)
        time.sleep(delay)
        time.sleep(delay)
        wd.find_element(By.ID, 'identifiant').send_keys(Keys.TAB)
        # # script = f'''identifant = document.getElementById('identifiant'); identifiant.setAttribute('type','hidden');
        # # identifiant.setAttribute('value',"youssef.atigui"); '''
        # wd.execute_script(script)

        # ## Saisie mot de pass
        time.sleep(delay)
        time.sleep(delay)
        wd.find_element(By.ID, 'secret_tmp').send_keys(mot_de_passe)
        time.sleep(delay)
        wd.find_element(By.ID, 'secret_tmp').send_keys(Keys.RETURN)
        time.sleep(delay)

    try:
        WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'ligneServiceHabilitation')))
    except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        logging.error("Except occurred", exc_info=True)
        messagebox.showinfo("Service Interrompu !", "Le service est indisponible\n pour l'instant")
        wd.quit()

    # ## Saisir service
    wd.find_element(By.ID, 'nomServiceChoisi').send_keys('0070100')  # FRP MEDOC DGE
    # # wd.find_element(By.ID, 'nomServiceChoisi').send_keys('6200100')
    # time.sleep(delay)
    wd.find_element(By.ID, 'nomServiceChoisi').send_keys(Keys.TAB)

    ## Saisir habilitation
    try:
        time.sleep(delay)
        wd.find_element(By.ID, 'habilitation').send_keys('1')
        time.sleep(delay)
        wd.find_element(By.ID, 'habilitation').send_keys(Keys.ENTER)
    except:
        progressbar_label.destroy()
        logging.error("Une Except est apparu", exc_info=True)
        WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
        messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
        messagebox.showinfo("Service Interrompu !", messages)
        wd.quit()

    ## Boucle sur le fichier selon le nombre de lignes indiquées
    j = 0
    while True:
        print(f"les données à la nouvelle ligne à la ligne 395: ", data)
        # time.sleep(delay)
        # wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys('3')
        # wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys(Keys.ENTER)
        # time.sleep(delay)

        # ## Création d'un Redevable
        # ## Arriver à la transaction 3-17
        # try:
        #     WebDriverWait(wd, 40).until(
        #         EC.presence_of_element_located((By.ID, 'bmenuxtableMenus:16:outputBmenuxBrmenx04LibelleLigneProposee')))
        #     wd.find_element(By.ID, 'bmenuxtableMenus:16:outputBmenuxBrmenx04LibelleLigneProposee').click()
        # except:
        #     progressbar_label.destroy()
        #     logging.error("Une Except est apparu", exc_info=True)
        #     messagebox.showinfo("Service Interrompu !", "La transaction création des oppositions ne semblent pas être "
        #                                                 "disponible. Veuillez tester manuellement avant de redémarrer "
        #                                                 "l'automate.")
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     break
        # print("N° de ligne à la ligne 416: ", j)
        num_of_secs = 540
        min, sec = divmod(num_of_secs * 3, 60)
        hour, min = divmod(min, 60)
        heure_de_fin = heure_demarrage + timedelta(hours=hour, minutes=min)
        heure_de_fin = heure_de_fin.strftime('%H:%M:%S')
        progressbar_label.destroy()
        progressbar_label = Label(tab6,
                                  text=f"Le travail est en cours: {pb['value']:.2f}%  ~  "
                                       f"L'opération de création de SATD sera terminé à  {heure_de_fin}")
        progressbar_label.place(x=250, y=label_y)
        tab6.update()
        #
        # ## Saisie numéro de Dossier
        # while True:
        #     print(f"numéro de dossier pour la ligne {j} à ligne 431: ", data[j][0])
        #     time.sleep(delay)
        #     WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))
        #     wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[j][0])
        #     wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(Keys.ENTER)
        #     errorMessages = ""
        #     print("messages d'erreur: ", errorMessages)
        #     time.sleep(delay)
        #     time.sleep(delay)
        #     time.sleep(delay)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
        #     errorMessages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
        #     print("messages d'erreur: ", errorMessages)
        #     messageDossierVerrouille = "DOSSIER DEJA UTILISE PAR UN AUTRE POSTE  - ATTENTE OU ABANDON - ".replace(" ",
        #                                                                                                           "")
        #
        #     errorMessagesIsPresent = errorMessages.replace(" ", "") == messageDossierVerrouille
        #     time.sleep(delay)
        #     time.sleep(delay)
        #     if errorMessagesIsPresent:
        #         messages = f"{errorMessages} \n Le dossier N°{data[j][0]} est ouvert par un autre agent ou verrouillé." \
        #                    f"\n Vous pouvez relancer le processus. Cette ligne sera exclu et pourra être relancer dans " \
        #                    f"45 minutes"
        #         messagebox.showinfo("Dossier verrouillé !", messages)
        #         data[j][15] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        #         data[j][16] = '\U0001F512'
        #         time.sleep(delay)
        #         j = j + 1
        #     else:
        #         break
        #     print("le N° de ligne est à la ligne 461 :", j)
        #
        # ## Saisie du choix Créer
        # print(f"les données à la nouvelle ligne {j} à la ligne 464: ", data[j])
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI')))
        #     wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys('C')
        #     wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys(Keys.TAB)
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     logging.error(
        #         "une except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException) est apparu en ligne 457",
        #         exc_info=True)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        #
        # ## Saisie du numéro de dossier créancier
        # try:
        #     time.sleep(delay)
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))
        #     # wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(numero_creancier_opposant)
        #     wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[j][1])
        #     wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(Keys.TAB)
        #     print("ligne 374: ok")
        #     print("le N° de ligne est à la ligne 605 :", j)  # print(data[i][1])
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     logging.error(
        #         "une except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException) est apparu",
        #         exc_info=True)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        #
        # ## Saisie de la suite
        # try:
        #     time.sleep(delay)
        #     time.sleep(delay)
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33gsuitYa33G002ReponseSuite')))
        #     wd.find_element(By.ID, 'inputB33gsuitYa33G002ReponseSuite').send_keys('S')
        #     wd.find_element(By.ID, 'inputB33gsuitYa33G002ReponseSuite').send_keys(Keys.TAB)
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     print("data", data[j])
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        #
        # ## SAISIE DES REFERENCES DE L'OPPOSITION
        # ## Transport de créance
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(
        #         EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance')))
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance').send_keys('N')
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance').send_keys(Keys.TAB)
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     print("data", data[j])
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        #
        # ## Saisie ATD
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GadtAdt')))
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GadtAdt').send_keys('O')
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GadtAdt').send_keys(Keys.TAB)
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     print("data", data[j])
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #
        # ## Saisie du crédit
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GcredCreditIs')))
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GcredCreditIs').send_keys('N')
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GcredCreditIs').send_keys(Keys.TAB)
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     print("data", data[j])
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        #
        # ## Saisie Empêchement
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GempEmpechement')))
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GempEmpechement').send_keys('N')
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GempEmpechement').send_keys(Keys.TAB)
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     print("data", data[j])
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        #
        # ## Saisie Montant
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GmtMontant')))
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GmtMontant').send_keys(data[j][2])
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GmtMontant').send_keys(Keys.TAB)
        #     # print(data[i][2])
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     print("data", data[j])
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        #
        ## Saisie de la Date d'Effet
        print(type(data[j][3]))
        if isinstance(data[j][3], str):
            date_d_effet = datetime.strptime(data[j][3], "%Y-%m-%d")
            print("ici c'est un string")
            print(date_d_effet.day)
        else:
            date_d_effet = data[j][3]
            print("ici ce n'est pas un string")
        # exit()
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour')))
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour').send_keys(date_d_effet.day)
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour').send_keys(Keys.TAB)
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     print("data", data[j])
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        #
        # ## Saisie du Mois d'Effet
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois')))
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois').send_keys(date_d_effet.month)
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois').send_keys(Keys.TAB)
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     print("data", data[j])
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        #
        # ## Saisie de l'Année d'Effet
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee')))
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee').send_keys(date_d_effet.year)
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee').send_keys(Keys.TAB)
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     print("data", data[j])
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        #
        # ## Saisie de la référence de jugement
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(
        #         EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite')))
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite').send_keys(data[j][4])
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite').send_keys(Keys.TAB)
        #     # print(data[i][4])
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     print("data", data[j])
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        #
        # ## Saisie de la date d'exécution de jugement
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(
        #         EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementJour')))
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementJour').send_keys(Keys.TAB)
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(
        #         EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementMois')))
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementMois').send_keys(Keys.TAB)
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(
        #         EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementAnnee')))
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementAnnee').send_keys(Keys.TAB)
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     print("data", data[j])
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        #
        # ## Saisie de la date de renouvellement
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(
        #         EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementJour')))
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementJour').send_keys(Keys.TAB)
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(
        #         EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementMois')))
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementMois').send_keys(Keys.TAB)
        #     time.sleep(delay)
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(
        #         EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee')))
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee').send_keys(Keys.TAB)
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     print("data", data[j])
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        #
        # ## Validation de la non saisie des dates
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(
        #         EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee')))
        #     wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee').send_keys(Keys.TAB)
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     print("data", data[j])
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        #
        # ## Validation de la suite
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(
        #         EC.presence_of_element_located((By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec')))
        #     wd.find_element(By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec').send_keys('S')
        #     wd.find_element(By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec').send_keys(Keys.TAB)
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     print("data", data[j])
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        #
        # ## Validation de la saisie de l'opposition
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(
        #         EC.presence_of_element_located((By.ID, 'inputB33gvlcrYa33GvalcValidationCreation')))
        #     wd.find_element(By.ID, 'inputB33gvlcrYa33GvalcValidationCreation').send_keys('O')
        #     wd.find_element(By.ID, 'inputB33gvlcrYa33GvalcValidationCreation').send_keys(Keys.TAB)
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     print("data", data[j])
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        #
        # ## Capture numéro d'opération
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'outputB33gnopeYa33GnopeNOpe')))
        #     numero_ope = wd.find_element(By.ID, 'outputB33gnopeYa33GnopeNOpe').text
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     print("data", data[j])
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        #
        # ## Saisie de la fin de la phase 1bis
        # try:
        #     time.sleep(delay)
        #     WebDriverWait(wd, 40).until(
        #         EC.presence_of_element_located((By.ID, 'inputB33gnouvYa33GnvopNouvelleOpposition')))
        #     wd.find_element(By.ID, 'inputB33gnouvYa33GnvopNouvelleOpposition').send_keys('N')
        #     wd.find_element(By.ID, 'inputB33gnouvYa33GnvopNouvelleOpposition').send_keys(Keys.TAB)
        # except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
        #     messagebox.showinfo("Service Interrompu !", message_service_interrompu)
        #     print("data", data[j])
        #     sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.quit()
        #     exit()
        # exit()
        # http: // medoc.appli.dgfip: 8140 / medocweb / presentation / md2oagt / ouverturesessionagent / ecran / ecOuvertureSessionAgent.jsf
        # Portail
        # applicatif
        # http: // medoc.appli.dgfip: 8140 / medocweb / index.xhtml
        ## -----------------Début de la phase 2-----------------------------------------------
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
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Création affaire service au code R17 "7055"
        # Saisie de la nature "AFF" pour debit 473-0
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('AFF')
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)
            print("pas 1 - ligne 799")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie du type de montant
        try:
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
            print("pas 2 - ligne 814")

        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie du montant X
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(
                data[j][7])
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(
                Keys.ENTER)
            print("pas 3 - ligne 868")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

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
            print("pas 4 - ligne 856")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie du numéro d'affaire
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03Nuaff1NumeroAffaire')))
            wd.find_element(By.ID, 'inputBcaff03Nuaff1NumeroAffaire').send_keys(data[j][5])
            print("numèro affaire:", data[j][5])
            print("pas 5 - ligne 873")
            time.sleep(delay)
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Confirmer le libelle de l'affaire
        try:
            time.sleep(delay)
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03LaffAffLibelleAffaire')))
            wd.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(Keys.ENTER)
            print("pas 6 - ligne 891")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisir le code R27 "7370"
        try:
            time.sleep(delay)
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi')))
            wd.find_element(By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi').send_keys(Keys.ENTER)
            print("ligne 832")
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01RMontantSaisi')))
            time.sleep(delay)
            print("ligne 836")
            wd.find_element(By.ID, 'inputBcaff01Bcaff01RMontantSaisi').send_keys(Keys.ENTER)
            time.sleep(delay)
            print("ligne 839")
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01SValidationOperateur')))
            wd.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys('O')
            wd.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys(Keys.ENTER)
            time.sleep(delay)
            # WebDriverWait(wd, 40).until(
            #     EC.presence_of_element_located((By.ID, 'inputBcaff12Bcaff121ValidationON')))
            # wd.find_element(By.ID, 'inputBcaff12Bcaff121ValidationON').send_keys('O')
            # wd.find_element(By.ID, 'inputBcaff12Bcaff121ValidationON').send_keys(Keys.ENTER)
            print("pas 7 - ligne 840")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisir le numéro du compte 477-0
        try:
            time.sleep(delay)
            time.sleep(delay)
            time.sleep(delay)
            time.sleep(delay)
            # credit_script = f''' let event = {enter}; numero_ligne = document.getElementById(
            # 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI'); numero_ligne.dispatchEvent(event); '''
            # wd.execute_script(credit_script)
            # input_list = wd.find_elements(By.ID,'bcaff01AffectationMvtsAUneAffairePanel')
            # print("liste des éléments : ",input_list)

            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI')))
            print("ok")
            time.sleep(delay)
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI').send_keys(Keys.ENTER)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx')))
            time.sleep(delay)
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx').send_keys('477-0')
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx').send_keys(Keys.ENTER)
            print("pas 8 - ligne 857")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisir la nature "AFF" pour crédit 477-0
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('AFF')
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)
            print("pas 9 - ligne 874")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie du type de montant
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
            print("pas 10 - ligne 890")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

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
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie la date comptable
        # Capture et réutilisation de la date journée comptable
        try:
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'PDATCPT_dateJourneeComptable')))
            djc_capture = wd.find_element(By.ID, 'PDATCPT_dateJourneeComptable').text
            djc = djc_capture.split('/')
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation').send_keys(djc[0])
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation').send_keys(Keys.ENTER)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation').send_keys(djc[1])
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation').send_keys(Keys.ENTER)
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation').send_keys(djc[2])
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation').send_keys(Keys.ENTER)
            # time.sleep(delay)
            # wd.find_element(By.XPATH, '//*[@id="repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation"]').send_keys(
            #     Keys.ENTER)
            print("pas 12 - ligne 936")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie du numéro d'affaire
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03Nuaff1NumeroAffaire')))
            wd.find_element(By.ID, 'inputBcaff03Nuaff1NumeroAffaire').send_keys(Keys.ENTER)
            print("pas 13 - ligne 951")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

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
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

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
            print("pas 15 - ligne 1138")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

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
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01SValidationOperateur')))
            wd.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys('O')
            print("pas 16 - ligne 1167")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Validation de la transaction
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcvim01Ycvim013ReponseOperateur')))
            wd.find_element(By.ID, 'inputBcvim01Ycvim013ReponseOperateur').send_keys('O')
            print("pas 17 - ligne 1032")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Création d'une liste temporaire avec numéro d'ordre de dépenses, le numéro de l'affaire créée et le numéro
        # de l'opération Le numéro de l'opération est divisé sur deux cellules dans MEDOC Cette liste sera finalement
        # collée comme ligne dans le fichier des donnees de sortie
        liste_temporaire_data = [str(data[j][0]), str(data[j][7])]  # FRP indice #0 dans liste_temporaire_data
        # #Montant indice #1 dans liste_temporaire_data

        # Numéro de l'ordre de dépense 1
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'outputBcvcs04Ycvcs022NoDepense')))
            # Numéro de l'ordre de dépense indice #2 dans liste_temporaire_data
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2')))
            liste_temporaire_data.append(wd.find_element(By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2').text)
            print("pas 18 - ligne 1056")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Pour afficher la suite
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvcs04Ycvcs028Reponse')))
            wd.find_element(By.ID, 'inputBcvcs04Ycvcs028Reponse').send_keys(Keys.ENTER)
            print("pas 19 - ligne 1070")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Numéro de l'affaire créée
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'outputBcvcs04Ycvcs021NumAffaireCreee')))
            # Numéro de l'affaire créée indice #3 dans liste_temporaire_data
            numero_affaire_creee = wd.find_element(By.ID, 'outputBcvcs04Nuaff1NumeroAffaire').text
            liste_temporaire_data.append(numero_affaire_creee)
            print("pas 20 - ligne 1087")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

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
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Numero de l'opération 1
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'labelBcvcs032')))
            # Numero de l'opération indice #4 dans liste_temporaire_data
            liste_temporaire_data.append(
                wd.find_element(By.ID, 'outputBcvcs03Nuopet1ErCarNuopeF2').text +
                wd.find_element(By.ID, 'outputBcvcs03Nuopes5DerniersCarNuope').text)
            print("pas 22 - ligne 1127")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Pour afficher la suite
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvcs03Ycvcs014DemandeSuite')))
            wd.find_element(By.ID, 'inputBcvcs03Ycvcs014DemandeSuite').send_keys('S')
            print("pas 23 - ligne 1152")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Fin de la transaction 21-2 et retour à la page d'accueil
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'barre_outils:image_f2')))
            wd.find_element(By.ID, 'barre_outils:image_f2').click()
            print("pas 24 - ligne 1274")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

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
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            progressbar_label.destroy()
            # WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            # messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messages = "Une erreur inattendu"
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Création affaire service au code R27 "8755"
        # Saisir la nature "AFF" pour debit 473-0
        # Saisir la nature "AFF" pour crédit 477-0
        try:
            time.sleep(delay)
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('AFF')
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)
            print("pas 26 - ligne 1199")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie du type de montant
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
            print("pas 27 - ligne 1215")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

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
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

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
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisir le numéro d'affaire créée précédemment
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03Nuaff1NumeroAffaire')))
            wd.find_element(By.ID, 'inputBcaff03Nuaff1NumeroAffaire').send_keys(liste_temporaire_data[3])
            print("pas 30 - ligne 1260")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Confirmer le libelle de l'affaire
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03LaffAffLibelleAffaire')))
            wd.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(Keys.ENTER)
            print("pas 31 - ligne 1276")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisir le code R27 "8755"
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi')))
            wd.find_element(By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi').send_keys(Keys.ENTER)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01RMontantSaisi')))
            wd.find_element(By.ID, 'inputBcaff01Bcaff01RMontantSaisi').send_keys(Keys.ENTER)
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01SValidationOperateur')))
            wd.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys('O')
            print("pas 32 - ligne 1298")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Répondre à la question "Soldez-vous l'affaire ?"
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff12Bcaff121ValidationON')))
            wd.find_element(By.ID, 'inputBcaff12Bcaff121ValidationON').send_keys('O')
            print("pas 33 - ligne 1314")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

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
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

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
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisir ENTREE pour type de montant
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
            print("pas 36 - ligne 1368")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

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
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

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
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Appuyer sur Entrer pour continuer
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBn4f3001Bn4F300116ZoneAcquisitionLibre')))
            wd.find_element(By.ID, 'inputBn4f3001Bn4F300116ZoneAcquisitionLibre').send_keys(Keys.ENTER)
            print("pas 39 - ligne 1435")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Validation de la transaction
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcvim01Ycvim013ReponseOperateur')))
            wd.find_element(By.ID, 'inputBcvim01Ycvim013ReponseOperateur').send_keys('O')
            print("pas 40 - ligne 1451")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Numéro de l'ordre de dépense 2
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2')))
            # Numero de l'ordre de depense indice #5 dans liste_temporaire_data
            liste_temporaire_data.append(wd.find_element(By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2').text)
            print("pas 41 - ligne 1452")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Pour afficher la suite
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcvcs04Ycvcs028Reponse')))
            wd.find_element(By.ID, 'inputBcvcs04Ycvcs028Reponse').send_keys(Keys.ENTER)
            print("pas 42 - ligne 1468")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Numéro de l'opération 2
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'labelBcvcs032')))
            # Numero de l'opération indice #6 dans liste_temporaire_data
            liste_temporaire_data.append(
                wd.find_element(By.ID, 'outputBcvcs03Nuopet1ErCarNuopeF2').text +
                wd.find_element(By.ID, 'outputBcvcs03Nuopes5DerniersCarNuope').text)
            print("pas 43 - ligne 1491")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Pour afficher la suite
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvcs03Ycvcs014DemandeSuite')))
            wd.find_element(By.ID, 'inputBcvcs03Ycvcs014DemandeSuite').send_keys('S')
            print("pas 44 - ligne 1506")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

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

        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisir le numéro d'affaire à partir des données d'entrées
        try:
            time.sleep(delay)
            time.sleep(delay)
            time.sleep(delay)
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
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisir le type de l'affaire "64"
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBrsdo03NasdoNatureSousDossier')))
            wd.find_element(By.ID, 'inputBrsdo03NasdoNatureSousDossier').send_keys('64')
            print("pas 48 - ligne 1593")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Récuperer le nouveau solde de l'affaire au code 1760 et enregistrer le sous indice #7 dans liste_temporaire_data
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'outputBraff01Yraff01YSoldeArticle')))
            wd.find_element(By.ID, 'outputBraff01Yraff01YSoldeArticle').text
            liste_temporaire_data.append(wd.find_element(By.ID, 'outputBraff01Yraff01YSoldeArticle').text)
            print("pas 49 - ligne 1576")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Récupérer le nom de l'entreprise à rembourser et enregistrer le dans une liste temporaire
        liste_tempo_nom_entreprise = [str(data[j][0])]
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
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Pour afficher la suite encore une fois en cas de besoin
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputYrval18wAcquisitionEspace')))
            wd.find_element(By.ID, 'inputYrval18wAcquisitionEspace').send_keys(Keys.ENTER)
            print("pas 52 - ligne 1617")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

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
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisir la nature "AFF" pour debit 473-0
        try:
            time.sleep(delay)
            WebDriverWait(wd, 30).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('AFF')
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)
            print("pas 55 - ligne 1661")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie du type de montant
        try:
            time.sleep(delay)
            WebDriverWait(wd, 10).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
            print("pas 56 - ligne 1676")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie du montant X
        montant = liste_temporaire_data[7].replace('+', '')
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
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie du numéro d'affaire créée précédemment
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03Nuaff1NumeroAffaire')))
            wd.find_element(By.ID, 'inputBcaff03Nuaff1NumeroAffaire').send_keys(data[j][5])
            print("pas 59 - ligne 1725")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Confirmer le libelle de l'affaire
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff03LaffAffLibelleAffaire')))
            wd.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(Keys.ENTER)
            print("pas 60 - ligne 1740")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

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
            print("pas 61 - ligne 1810")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Répondre à la question "Soldez-vous l'affaire ?"
        try:
            time.sleep(delay)
            WebDriverWait(wd, 60).until(
                EC.presence_of_element_located((By.ID, 'inputBcaff12Bcaff121ValidationON')))
            wd.find_element(By.ID, 'inputBcaff12Bcaff121ValidationON').send_keys('O')
            print("pas 62 - ligne 1826")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Valider CREDIT
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI').send_keys(Keys.ENTER)
            print("pas 63 - ligne 1842")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisir le numéro du compte 512-96
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx').send_keys('512-96')
            print("pas 64 - ligne 1858")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie de la nature "VIRT"
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('VIRT')
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)
            print("pas 65 - ligne 1875")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie du type de montant
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
            print("pas 66 - ligne 1838")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie du montant X
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(
                liste_temporaire_data[7])
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(Keys.ENTER)
            print("pas 67 - ligne 1855")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie de la date du jour comptable
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation').send_keys(djc[0])
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation').send_keys(Keys.ENTER)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation').send_keys(djc[1])
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation').send_keys(Keys.ENTER)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation')))
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation').send_keys(djc[2])
            wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation').send_keys(Keys.ENTER)
            print("pas 68 - ligne 2089")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie du numéro de dossier
        try:
            time.sleep(delay)
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))
            wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[j][0])
            if wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').is_displayed:
                wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[j][0])
            print("pas 69 - ligne 2277")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Continuer en cas d'existence de ce message : ATTENTION - OPPOSITION POUR CE DOSSIER
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_all_elements_located((By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON')))
            wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').send_keys('O')
            print("pas 70 - ligne 1988")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

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
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Libelle du virement emis
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_all_elements_located((By.ID, 'repeatBcimp01:1:inputBcrib01IbanlibLibelleVirementEmis')))
            wd.find_element(By.ID, 'repeatBcimp01:1:inputBcrib01IbanlibLibelleVirementEmis').send_keys(
                str(data[j][12]) + "/ RCTVA")
            wd.find_element(By.ID, 'repeatBcimp01:1:inputBcrib01IbanlibLibelleVirementEmis').send_keys(Keys.ENTER)
            print("pas 72 - ligne 1946")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Répondre à la question "Voulez-vous valider ?"
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcvim01Ycvim013ReponseOperateur')))
            wd.find_element(By.ID, 'inputBcvim01Ycvim013ReponseOperateur').send_keys('O')
            print("pas 73 - ligne 1961")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Visualisation du Numero de l'ordre de dépense
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2')))
            # Numero de l'ordre de dépense indice #8 dans liste_temporaire_data
            liste_temporaire_data.append(wd.find_element(By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2').text)
            print("pas 73 - ligne 1977")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Pour afficher la suite
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBcvcs04Ycvcs028Reponse')))
            wd.find_element(By.ID, 'inputBcvcs04Ycvcs028Reponse').send_keys(Keys.ENTER)
            print("pas 74 - ligne 1992")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Numero de l'opération
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'labelBcvcs032')))
            # Numero de l'opération indice #9 dans liste_temporaire_data
            liste_temporaire_data.append(
                wd.find_element(By.ID, 'outputBcvcs03Nuopet1ErCarNuopeF2').text +
                wd.find_element(By.ID, 'outputBcvcs03Nuopes5DerniersCarNuope').text)
            print("pas 75 - ligne 2082")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Pour afficher la suite
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvcs03Ycvcs014DemandeSuite')))
            wd.find_element(By.ID, 'inputBcvcs03Ycvcs014DemandeSuite').send_keys('S')
            print("pas 76 - ligne 2097")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Fin de la transaction 21-2 et retour à la page d'accueil
        time.sleep(delay)
        WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'barre_outils:image_f2')))
        wd.find_element(By.ID, 'barre_outils:image_f2').click()

        # Saisir la transaction 3-1-7
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi')))
            wd.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('3')
            time.sleep(delay)
            WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi')))
            wd.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('1')
            time.sleep(delay)
            WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx062ECaractere')))
            wd.find_element(By.ID, 'inputBmenuxBrmenx062ECaractere').send_keys('7')
            print("pas 77 - ligne 2123")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Création affaire service au code R17 "7055"
        # Saisir le numéro du dossier
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))
            wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[j][0])
            print("pas 78 - ligne 2139")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisir M pour mise à jour
        try:
            time.sleep(delay)
            WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'labelB33gmenu0')))
            wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys("M")
            print("pas 79 - ligne 2154")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisir le numéro FRP de l'opposant
        try:
            time.sleep(delay)
            wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[j][1])
            print("pas 80 - ligne 2168")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Taper sur la touche «Entrée» jusqu’à la case de saisie
        # Titre
        try:
            time.sleep(delay)
            WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginfoYa33Gtit2Titre')))
            wd.find_element(By.ID, 'inputB33ginfoYa33Gtit2Titre').send_keys(Keys.ENTER)
            # Profession/activité
            time.sleep(delay)
            WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginfoYa33GprofProfession')))
            wd.find_element(By.ID, 'inputB33ginfoYa33GprofProfession').send_keys(Keys.ENTER)
            # Adresse No
            time.sleep(delay)
            WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf1Ya33GnoaNAdresse')))
            wd.find_element(By.ID, 'inputB33ginf1Ya33GnoaNAdresse').send_keys(Keys.ENTER)
            # Adresse B/T
            time.sleep(delay)
            WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf1Ya33GcplComplementBT')))
            wd.find_element(By.ID, 'inputB33ginf1Ya33GcplComplementBT').send_keys(Keys.ENTER)
            # Adresse VOIE
            time.sleep(delay)
            WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf1Ya33GvoieVoie')))
            wd.find_element(By.ID, 'inputB33ginf1Ya33GvoieVoie').send_keys(Keys.ENTER)
            # Adresse COMMUNE
            time.sleep(delay)
            WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf1Ya33GcomCommune')))
            wd.find_element(By.ID, 'inputB33ginf1Ya33GcomCommune').send_keys(Keys.ENTER)
            # Adresse COMPL
            time.sleep(delay)
            WebDriverWait(wd, 30).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf1Ya33GcpaComplementAdresse')))
            wd.find_element(By.ID, 'inputB33ginf1Ya33GcpaComplementAdresse').send_keys(Keys.ENTER)
            # Adresse CODE POSTAL
            time.sleep(delay)
            WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf1Ya33GcpoCodePostal')))
            wd.find_element(By.ID, 'inputB33ginf1Ya33GcpoCodePostal').send_keys(Keys.ENTER)
            # Adresse BUREAU
            time.sleep(delay)
            WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf1Ya33GbudiBureau')))
            wd.find_element(By.ID, 'inputB33ginf1Ya33GbudiBureau').send_keys(Keys.ENTER)
            print("pas 81 - ligne 2208")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisir S pour la validation
        try:
            time.sleep(delay)
            WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33gsuvlYa33G009Reponse')))
            wd.find_element(By.ID, 'inputB33gsuvlYa33G009Reponse').send_keys("S")
            print("pas 82 - ligne 2222")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Taper sur la touche «Entrée» jusqu’à la case de saisie
        # Transport de creance O/N
        try:
            time.sleep(delay)
            WebDriverWait(wd, 30).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance').send_keys(Keys.ENTER)

            # ATD, Saisies O/N
            time.sleep(delay)
            WebDriverWait(wd, 30).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GadtAdt')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GadtAdt').send_keys(Keys.ENTER)

            # Crédit O/N
            time.sleep(delay)
            WebDriverWait(wd, 30).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GcredCreditIs')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GcredCreditIs').send_keys(Keys.ENTER)
            # Empechements O/N
            time.sleep(delay)
            WebDriverWait(wd, 30).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GempEmpechement')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GempEmpechement').send_keys(Keys.ENTER)
            # Montant
            time.sleep(delay)
            WebDriverWait(wd, 30).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GmtMontant')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GmtMontant').send_keys(Keys.ENTER)
            # Date d'effet
            # Jour
            time.sleep(delay)
            WebDriverWait(wd, 30).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour').send_keys(Keys.ENTER)
            # Mois
            time.sleep(delay)
            WebDriverWait(wd, 30).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois').send_keys(Keys.ENTER)
            # Année
            time.sleep(delay)
            WebDriverWait(wd, 30).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee').send_keys(Keys.ENTER)
            # Ref jugt validite
            time.sleep(delay)
            WebDriverWait(wd, 30).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite')))
            # wd.find_element(By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite').send_keys(data[j][4])
            wd.find_element(By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite').send_keys(Keys.ENTER)
            # date d'execution jugt
            # Jour
            time.sleep(delay)
            WebDriverWait(wd, 30).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementJour')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementJour').send_keys(Keys.ENTER)
            # Mois
            time.sleep(delay)
            WebDriverWait(wd, 30).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementMois')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementMois').send_keys(Keys.ENTER)
            # Année
            time.sleep(delay)
            WebDriverWait(wd, 30).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementAnnee')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementAnnee').send_keys(Keys.ENTER)
            print("pas 83- ligne 2274")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Date de renouvellement
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementJour')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementJour').send_keys(Keys.ENTER)
            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementMois')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementMois').send_keys(Keys.ENTER)
            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee')))
            time.sleep(delay)
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee').send_keys(Keys.ENTER)
            if wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementJour').is_displayed():
                wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementJour').send_keys(Keys.TAB)
                wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementMois').send_keys(Keys.TAB)
                wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee').send_keys(Keys.TAB)
                time.sleep(delay)
            # elif wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').is_displayed():
            #     time.sleep(delay)
            #     wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').send_keys('O')
            else:
                pass
            print("pas 84 - ligne 2505")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Message informatif
        # Pas de date d'exécution du jugement saisie, souhaitez-vous continuer ?
        try:
            wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').is_displayed()
            time.sleep(delay)
            wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').send_keys('O')
        except:
            pass

        # Saisir S pour passer à l'écran suivant
        try:
            time.sleep(delay)
            WebDriverWait(wd, 30).until(
                EC.presence_of_element_located((By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec')))
            wd.find_element(By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec').send_keys('S')
            print("pas 85 - ligne 2521")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisir REF main levée
        try:
            time.sleep(delay)
            WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GrfmlRefMainlevee')))
            time.sleep(delay)
            numero_ope = data[j][14]
            wd.find_element(By.ID, 'inputB33ginf3Ya33GrfmlRefMainlevee').send_keys(f"{numero_ope} du {djc_capture}")
            time.sleep(delay)
            wd.find_element(By.ID, 'inputB33ginf3Ya33GrfmlRefMainlevee').send_keys(Keys.ENTER)
            print("pas 86 - ligne 2538")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisir date de main levée
        # Capture et réutilisation de la date journée comptable
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GdtmlDateMainleveeJour')))
            wd.find_element(By.ID, 'inputB33ginf3Ya33GdtmlDateMainleveeJour').send_keys(djc[0])
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GdtmlDateMainleveeMois')))
            wd.find_element(By.ID, 'inputB33ginf3Ya33GdtmlDateMainleveeMois').send_keys(djc[1])
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GdtmlDateMainleveeAnnee')))
            wd.find_element(By.ID, 'inputB33ginf3Ya33GdtmlDateMainleveeAnnee').send_keys(djc[2])
            wd.find_element(By.ID, 'inputB33ginf3Ya33GdtmlDateMainleveeAnnee').send_keys(Keys.ENTER)
            print("pas 87 - ligne 2565")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisir type de main levée
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GtpmlTypeMainlevee')))
            wd.find_element(By.ID, 'inputB33ginf3Ya33GtpmlTypeMainlevee').send_keys('TOTALE')
            time.sleep(delay)
            wd.find_element(By.ID, 'inputB33ginf3Ya33GtpmlTypeMainlevee').send_keys(Keys.ENTER)
            print("pas 87 - ligne 2582")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # REF JUGT NULLITE
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GrfnlRefJugtNullite')))
            wd.find_element(By.ID, 'inputB33ginf3Ya33GrfnlRefJugtNullite').send_keys(Keys.ENTER)
            print("pas 88 - ligne 2597")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # DATE JUGT NULLITE
        # Jour
        try:
            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GdtnlDateJugtNulliteJour')))
            wd.find_element(By.ID, 'inputB33ginf3Ya33GdtnlDateJugtNulliteJour').send_keys(Keys.ENTER)
            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GdtnlDateJugtNulliteMois')))
            wd.find_element(By.ID, 'inputB33ginf3Ya33GdtnlDateJugtNulliteMois').send_keys(Keys.ENTER)
            time.sleep(delay)
            WebDriverWait(wd, 20).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GdtnlDateJugtNulliteAnnee')))
            wd.find_element(By.ID, 'inputB33ginf3Ya33GdtnlDateJugtNulliteAnnee').send_keys(Keys.ENTER)
            print("pas 89 - ligne 2622")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisir O pour valider (Voulez-vous valider ?)
        try:
            time.sleep(delay)
            validation_script = \
                f'''document.getElementById('inputB33gvlcrYa33GvalcValidationCreation').setAttribute('value','O');'''
            wd.execute_script(validation_script)
            # WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33gvlcrYa33GvalcValidationCreation')))
            time.sleep(delay)
            wd.find_element(By.ID, 'inputB33gvlcrYa33GvalcValidationCreation').send_keys(Keys.ENTER)
            print("pas 90 - ligne 2641")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Création d'une liste temporaire avec numéro d'ordre de dépenses, le numéro de l'affaire créée et le numéro de l'opération
        # Le numéro de l'opération est divisé sur deux cellules dans MEDOC
        # Cette liste sera finalement collée comme ligne dans le fichier des donnees de sortie
        liste_temporaire_data = [str(data[j][0]), str(data[j][1]), str(data[j][3])]

        # Numero de l'opération
        try:
            time.sleep(delay)
            WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'labelB33gnope0')))
            # Numero de l'ordre de depense indice #3 dans liste_temporaire_data
            liste_temporaire_data.append(wd.find_element(By.ID, 'outputB33gnopeYa33GnopeNOpe').text)
            print("pas 91 - ligne 2662")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # fin de transaction :
        try:
            time.sleep(delay)
            fin_transaction_script = \
                f'''document.getElementById('inputB33gcrmdYa33G012Reponse').setAttribute('value','N');'''
            wd.execute_script(fin_transaction_script)
            print("pas 92 - ligne 2678")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            print("data", data[j])
            sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Retour au menu
        WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        wd.find_element(By.ID, 'barre_outils:touche_f2').click()

        # convertir les colonnes numériques de la liste data en entier

        ## Marquage tâche faîte dans le fichier
        match os.path.isfile(filepath1):
            case True:
                data[j][13] = f"{numero_ope} du {djc_capture}"
                # data[j][14] = numero_ope
                data[j][16] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                data[j][15] = success
                print("inscription des données dans la liste ligne 2991", data)
            case False:
                data[j][13] = f"{numero_ope} du {djc_capture}"
                data[j][3] = str(date_d_effet.strftime('%Y-%m-%d'))
                # data[j][14] = numero_ope
                data[j][16] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                data[j][15] = success
                print("inscription des données ligne 2147", data)
        print("le N° de ligne est à la ligne 2148:", j)

        ## Incrementation ProgressBar
        pb['value'] += 90 / nb_ligne
        progressbar_label.destroy()
        tab6.update()
        progressbar_label = Label(tab6,
                                  text=f"Le travail est en cours: {pb['value']:.2f}%  ~  "
                                       f"L'opération de création de SATD sera terminé à  {heure_de_fin}")
        progressbar_label.place(x=250, y=label_y)
        pb.update()
        tab6.update()
        print("les données à la nouvelle ligne : ", data[j])
        data_ods = data
        sheet_name = "Feuille1"
        data_ods.insert(0, columns_sortie)
        print("les nouvelles data 2725: \n", data)

        if not os.path.exists(sortie_repertoire):
            os.makedirs(sortie_repertoire)
        if os.path.exists(sortie_repertoire + '/' + fichier_de_Sortie):
            print("old_data 2730 : \n", old_data)
            del data_ods[0]
            print("data sans les entêtes (ligne 2732)", data_ods)
            if not old_data:
                data_ods.insert(0, columns_sortie)
            else:
                numpyData = np.append(data_ods, old_data, axis=0)
                # data_ods = list(numpyData)
                data_ods = numpyData.tolist()
                data_ods.insert(0, columns_sortie)
            print("listData : \n", data_ods)
            os.remove(sortie_repertoire + '/' + fichier_de_Sortie)
        else:
            print("old_data 2953: \n", old_data)
            del data_ods[0]
            print("data sans les entêtes (ligne 2745)", data_ods)
            if not old_data:
                data_ods = data_ods
            else:
                numpyData = np.append(data_ods, old_data, axis=0)
                data_ods = numpyData.tolist()
            data_ods.insert(0, columns_sortie)
            print("data_ods : \n", data_ods)
            print("type data_ods : \n", type(data_ods))
            print("pas 93 - ligne 2754 - Fin de boucle")
        save_data(saved_file, data_ods)
        mouse.position = (10, 20)
        if j + 1 < nb_ligne:
            j += 1
        else:
            break
        del data[0]
        print("ligne 2973: ", data)
    data_df = pd.DataFrame(data_ods)

    print("le dataframe : ", data_df)

    try:
        time.sleep(delay)
        time.sleep(delay)
        time.sleep(delay)
        tabControl.add(tab4, text='Liste SATD en sortie de traitement')
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
    messagebox.showinfo("Données Manquante", "Le travail est maintenant fini! A bientôt")
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
        messagebox.showinfo("SATD", 'Votre fichier contient ' + str(nb_ligne) + ' ligne' + s + '.')
        print('Votre fichier contient ' + str(nb_ligne) + ' ligne' + s + '.')
    fichier_de_Sortie = 'donnees_sortie' + datetime.now().strftime('_%Y-%m-%d') + '.ods'
    filepath1 = source_rep + '/sorties_SATD/sorties_SATD' + datetime.now().strftime(
        '_%Y-%m-%d') + '/' + fichier_de_Sortie
    print(os.path.isfile(filepath1))
    if os.path.isfile(filepath1):
        f = open(filepath1, "a")
        f.write("")
        f.close()
        df1 = pd.read_excel(filepath1, dtype={'a': str, 'b': str, 'c': np.int64, 'd': np.datetime64, 'e': str
            , 'f': np.int64, 'g': np.int64, 'h': np.int64, 'i': str, 'j': np.int64, 'k': np.int64, 'l': np.int64,
                                              'm': np.int64, 'n': str}, engine='odf').fillna(0)
        column1 = df1.columns[6]
        print("le dataframe des anciennes données : \n", df1)
        print("----------------------------------------------------------------------------")
        nb_ligne1 = df1.shape[0]
        s = 's' if nb_ligne1 > 1 else ''
        sub_df1 = df1[df1['Dossiers traités'] == success]
        print("le dataframe contenant les lignes déjà faites: \n", sub_df1)
        print("----------------------------------------------------------------------------")
        if len(sub_df1) - len(df) != 0:
            response = messagebox.askyesno(
                "Automate SATD", "Le fichier a déjà été traité par l'automate, à l'Except d'une ou "
                                 f"plusieurs SATD identifiée(s) par les symboles \"{success}, {vide}\""
                                 "ou M en colonne \"Dossiers traités\".")
            try:
                time.sleep(2)
                tab5 = Frame(tabControl, bg='#E3EBD0')
                tabControl.add(tab5, text='liste des SATD déjà effectuées')
                sub_df2 = df1[(df1['Dossiers traités'] == success) | (df1['Dossiers traités'] == vide) | (
                        df1['Dossiers traités'] == 'M')]
                print("liste des SATD déjà effectuées", sub_df2)
                # df1['Date d’effet = date réception SATD'] = df['Date d’effet = date réception SATD'].dt.strftime(
                #     '%d-%m-%Y')
                table = Table(tab5, dataframe=sub_df2, read_only=True, index=FALSE)
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
        colonne_13 = "Numéro et date de l'opération de dépense effectuée dans Médoc pour paiement du poste " \
                     "comptable RNF ayant émis la SATD "
        colonne_14 = "Dossiers traités"
        if df.drop(columns=[colonne_13, colonne_14]).loc[i].isnull().any():
            message = "la ligne {} du tableau comporte une ou plusieurs données obligatoires manquantes.\n Cette " \
                      "ligne ne sera pas traitée et sera marquée dans la colonne \"Dossiers traités\" par le symbole " \
                      "\"∅\". \n Vous pouvez renseigner les champs manquants avant de lancer l'automate.".format(i + 1)
            messagebox.showwarning("Données manquantes", message)
        elif df["N°affaire code 1760"].duplicated().loc[i] or df["Numéro de facture Chorus"].duplicated().loc[
            i]:
            message = "la ligne {} du tableau comporte ne comporte pas d'identifiant unique dans la colonne « Réf " \
                      "jugement validité = réf SATD » et ou la colonne « N°affaire code 1760 ». Cette ligne doit être " \
                      "traitée manuellement et sera marquée dans la colonne \"Dossiers traités\" par le symbole " \
                      "\"M\".".format(i + 1)
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
# TODO :enlever le code cadenas vérrouillé
lexique = "Précisions sur les symboles affichés en colonne \"Dossiers traités\" du fichier de sortie :\n" \
          "\n● Le symbole \"\u2713\" indique que la SATD a été traitée jusqu'à la mainlevée." \
          "\n● Le symbole \"∅\" indique qu'une ou plusieurs données obligatoires sont manquantes sur la ligne, ce qui " \
          "ne permet pas de traiter la SATD. Il convient de compléter la ou les données manquantes avant d'exécuter de" \
          " nouveau l'automate pour traiter la SATD concernée." \
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
entry5 = Entry(tab6, textvariable=EnterTable5,show="*",justify='center')
entry5.place(x=600, y=70)
def toggle_password():
    if entry5.cget('show') == '':
        time.sleep(10)
        entry5.config(show='*')
        button_font.config(overstrike = 1)
    else:
        entry5.config(show='')
button_font = font.Font(family='Tahoma',size=12)
show_password_btn = Button(tab6, text='👁',font=button_font ,justify=CENTER,command=toggle_password)
show_password_btn.place(x=730, y=65)
button2 = Button(tab6, bg="#CEDDDE", text='Choisir le fichier d\'entrée', command=open_file)
button2.place(x=paramx + 240, y=paramy - 30)
label_path6 = Label(tab6)
label_path6.place(x=paramx + 490, y=paramy - 30)

wd_button = Button(tab6, bg="#C7DDC5", text='Créer les SATD sans visualisation des transactions',
                   command=lambda: create_opposition(headless=True))
wd_button.place(x=paramx + 240, y=paramy + 100)
creerOpposition = Button(tab6, bg="#9FCDA8", text='Créer les SATD avec visualisation des transactions',
                         command=lambda: create_opposition(headless=False))
creerOpposition.place(x=paramx + 240, y=paramy + 150)

Interface.mainloop()
