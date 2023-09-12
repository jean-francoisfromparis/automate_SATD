import calendar
import gc
import locale
import os
import re
import shutil
from sys import exit
import time
import glob
from typing import List

import PyPDF2
from datetime import datetime, date, timedelta
from pathlib import Path
from tkinter import *
from tkinter import filedialog, messagebox, ttk, font
from tkinter.ttk import Progressbar

import dateparser
import pandas as pd
from PIL import Image, ImageTk
from pandastable import Table
from pyexcel_ods import save_data
from pynput.keyboard import Controller
from selenium import webdriver
from selenium.common import TimeoutException, StaleElementReferenceException, \
    ElementNotInteractableException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from zipfile import ZipFile
from _utils.save_file import Saved_file
from _utils.telecharger import Telecharger_fichier

mouse = Controller()
success = '✓'
vide = "∅"
delay = 4


def __init__(self, progress):
    self.progress = progress
    global delay


# Fonction pour retrouver le chemin d'accès
# def resource_path(relative_path):
#     try:
#         base_path = sys._MEIPASS
#     except Exception:
#         base_path = os.path.dirname(__file__)
#     return os.path.join(base_path, relative_path)


def create_opposition(headless):
    gc.collect()
    message_service_interrompu = "\nLa qualité de la connexion ne permet pas un bon fonctionnement de l'automate. " \
                                 "Veuillez essayer ultérieurement ! "

    # Initialisation de la methode de sauvegarde en cas d'erreurs
    sav = Saved_file()
    global success
    global vide

    # Etablissement du progressBar

    pb = progressbar(tab6)
    progressbar_label = Label(tab6, text=f"Le travail commence. L'automate se connecte...")
    label_y = 390
    progressbar_label.place(x=250, y=label_y)
    tab6.update()

    time.sleep(delay)

    # Prend la ligne du fichier depuis laquelle commencer à lire
    # while True:
    #     line = EnterTable2.get()
    #     if line.isnumeric():  ##vérifie que ça soit un numéro
    #         line = int(line)  ##ajuste l'indice
    #         break
    #     else:
    #         messagebox.showerror("Erreur de saisie", 'Saisie incorrecte, réessayez')
    #         exit()

    # Combien de lignes du fichier traiter
    # while True:
    #     line_amount = EnterTable3.get()
    #     if line_amount.isnumeric():
    #         line_amount = int(line_amount)
    #         break
    #     else:
    #         messagebox.showerror("Erreur de saisie", 'Saisie incorrecte, réessayez')
    #         exit()

    # Prend les données depuis le fichier, crée une liste de listes (ou "array"), oú chaque liste est
    # une ligne du fichier Calc. Il faut faire ça parce que pyxcel_ods prend les données sous forme
    # de dictionnaire.
    entree_df = pd.read_excel(File_path)
    entree_df.drop(entree_df.columns[[13, 14]], axis=1, inplace=True)
    source_rep = os.getcwd()
    fichier_de_sortie = 'donnees_sortie_phase1' + datetime.now().strftime('_%Y-%m-%d') + '.ods'
    repertoire_de_sortie = source_rep + '\donnees_sortie_phase1\donnees_sortie_phase1' + datetime.now().strftime(
        '_%Y-%m-%d')
    chemin_fichier_de_sortie = repertoire_de_sortie + '\\' + fichier_de_sortie
    print("---------------------------Récupération des données-------------------------------------------------")
    # Vérification de l'existence d'un fichier de sortie à la date du jour
    columns = list(entree_df)
    columns_sortie = columns + ["N°Opération Phase 1", "Date Opération phase 1", "Dossiers traités"]
    sortie_df = pd.DataFrame(columns=columns_sortie)
    if not os.path.exists(repertoire_de_sortie):
        os.makedirs(repertoire_de_sortie)
    elif not os.path.exists(chemin_fichier_de_sortie):
        sortie_df = pd.DataFrame(columns=columns_sortie)
    else:
        sortie_df = pd.read_excel(chemin_fichier_de_sortie)
    print("Le fichier d'entrée contient " + str(entree_df.shape[0]) + " lignes.")
    print("Le fichier de sortie contient " + str(
        sortie_df.shape[0]) + "Lignes de la précédente opération.")
    print(entree_df)
    # Création des colonnes afin de comparer les dataframes des données d'entrée avec les dataframes de données de
    # sortie
    entree_df["N° dossier FRP opposé"] = entree_df["N° dossier FRP opposé"].astype(
        str, errors='ignore')
    entree_df["N° dossier FRP opposant"] = entree_df[
        "N° dossier FRP opposant"].astype(
        str, errors='ignore')
    entree_df["Montant opposition"] = entree_df[
        "Montant opposition"].astype(
        int, errors='ignore')
    entree_df["Numéro de facture Chorus"] = entree_df[
        "Numéro de facture Chorus"].astype(
        str, errors='ignore')
    entree_df["N°affaire code 1760"] = \
        entree_df["N°affaire code 1760"].astype(
            int, errors='ignore')
    entree_df["Montant de l’affaire au code 1760"] = \
        entree_df["Montant de l’affaire au code 1760"].astype(
            int, errors='ignore')
    entree_df["Montant à créer en « affaire service » au code 7055"] = \
        entree_df["Montant à créer en « affaire service » au code 7055"].astype(
            int, errors='ignore')
    entree_df["Identification du bénéficiaire de la dépense"] = entree_df[
        "Identification du bénéficiaire de la dépense"].astype(
        str, errors='ignore')
    entree_df["Codique du service bénéficiaire"] = \
        entree_df["Codique du service bénéficiaire"].astype(str)
    entree_df["RANG RIB pour le remboursement du service bénéficiaire"] = \
        entree_df["RANG RIB pour le remboursement du service bénéficiaire"].astype(
            int, errors='ignore')
    entree_df["RANG RIB pour le remboursement à la société "] = \
        entree_df["RANG RIB pour le remboursement à la société "].astype(
            int, errors='ignore')
    entree_df["SIREN du redevable pour le libellé du virement pour la société"] = \
        entree_df["SIREN du redevable pour le libellé du virement pour la société"].astype(
            int, errors='ignore')
    entree_df.insert(13, "N°Opération Phase 1", "0", allow_duplicates=False)
    entree_df["N°Opération Phase 1"] = entree_df[
        "N°Opération Phase 1"].astype(str)
    entree_df.insert(14, "Date Opération phase 1", "0", allow_duplicates=False)
    entree_df["Date Opération phase 1"] = entree_df[
        "Date Opération phase 1"].astype(str)
    entree_df["Dossiers traités"] = '0'
    entree_df["Dossiers traités"] = entree_df[
        "Dossiers traités"].astype(str)

    sortie_df["N° dossier FRP opposé"] = sortie_df["N° dossier FRP opposé"].astype(
        str, errors='ignore')
    sortie_df["N° dossier FRP opposant"] = sortie_df[
        "N° dossier FRP opposant"].astype(
        str, errors='ignore')
    sortie_df["Montant opposition"] = sortie_df[
        "Montant opposition"].astype(
        int, errors='ignore')
    sortie_df["N°affaire code 1760"] = \
        sortie_df["N°affaire code 1760"].astype(
            int, errors='ignore')
    sortie_df["Montant de l’affaire au code 1760"] = \
        sortie_df["Montant de l’affaire au code 1760"].astype(
            int, errors='ignore')
    sortie_df["Montant à créer en « affaire service » au code 7055"] = \
        sortie_df["Montant à créer en « affaire service » au code 7055"].astype(
            int, errors='ignore')
    sortie_df["Identification du bénéficiaire de la dépense"] = sortie_df[
        "Identification du bénéficiaire de la dépense"].astype(
        str, errors='ignore')
    sortie_df["Codique du service bénéficiaire"] = \
        sortie_df["Codique du service bénéficiaire"].astype(str)
    sortie_df["RANG RIB pour le remboursement du service bénéficiaire"] = \
        sortie_df["RANG RIB pour le remboursement du service bénéficiaire"].astype(
            int, errors='ignore')
    sortie_df["RANG RIB pour le remboursement à la société "] = \
        sortie_df["RANG RIB pour le remboursement à la société "].astype(
            int, errors='ignore')
    sortie_df["SIREN du redevable pour le libellé du virement pour la société"] = \
        sortie_df["SIREN du redevable pour le libellé du virement pour la société"].astype(
            int, errors='ignore')
    sortie_df["N°Opération Phase 1"] = sortie_df[
        "N°Opération Phase 1"].astype(str)
    sortie_df["Date Opération phase 1"] = sortie_df[
        "Date Opération phase 1"].astype(str)
    sortie_df["Dossiers traités"] = sortie_df[
        "Dossiers traités"].astype(str)
    # Vérification que le montant de l'affaire est bien strictement supérieur au montant de l'affaire à créer
    entree_df['comparaison'] = entree_df.apply(
        lambda x: True if x[6] <= x[7] else False, axis=1)
    ligne_incomplete = list()
    nb_ligne = entree_df.shape[0]
    print("ligne 120 \n", entree_df['comparaison'])
    for i in range(nb_ligne):
        if entree_df.drop(columns=['comparaison']).loc[i].isnull().any() or \
                entree_df["Date d’effet = date réception SATD"].loc[i] == 'NaT':
            ligne_incomplete.append(vide)
        elif entree_df['comparaison'].loc[i] or \
                entree_df["Numéro de facture Chorus"].duplicated().loc[i]:
            ligne_incomplete.append("M")
        else:
            ligne_incomplete.append('0')
    entree_df["Dossiers traités"] = ligne_incomplete
    print("ligne incomplete : ", ligne_incomplete)
    entree_df.drop(["comparaison"], axis=1, inplace=True)
    # Conservation des données déjà traitées
    sortie_traitee_df = sortie_df[
        (sortie_df["Dossiers traités"] == success) | (sortie_df["Dossiers traités"] == 'M') | (
                sortie_df["Dossiers traités"] != '0')]
    # Filtrage des données à traitées
    entree_df["filter"] = entree_df["Numéro de facture Chorus"].isin(sortie_traitee_df["Numéro de facture Chorus"])
    entree_df = entree_df[(entree_df["filter"] == False)]
    entree_df.drop(["filter"], axis=1, inplace=True)
    sortie_df["filter"] = sortie_df["N°affaire code 1760"].isin(entree_df["N°affaire code 1760"])
    sortie_df = sortie_df[(sortie_df["filter"] == False)]
    sortie_df.drop(["filter"], axis=1, inplace=True)
    print(sortie_traitee_df)
    print(entree_df)
    data = entree_df.values.tolist()
    sortie = sortie_df.values.tolist()
    print("les données d'entrées sont :", data)
    print("les données de sorties sont :", sortie)

    # conversion des données de date au format date :
    if sortie:
        for k in range(len(sortie)):
            if isinstance(sortie[k][3], str) or sortie[k][3] == 0:
                print("Ligne 166 :", type(sortie[k][3]))
                pass
            else:
                sortie[k][3] = sortie[k][3].strftime('%Y-%m-%d')
                print("Ligne 170:", type(sortie[k][3]))
    # sauvegarde des anciennes données avec les nouvelles données qui ne doivent pas être traitées
    sortie.insert(0, columns_sortie)
    save_data(chemin_fichier_de_sortie, sortie)
    # Remplacement des espaces dans la liste
    nb_ligne = len(data)
    for i in range(nb_ligne):
        data[i][0] = str(data[i][0]).replace(" ", "")
        data[i][1] = str(data[i][1]).replace(" ", "")
        data[i][9] = str(data[i][9]).replace(" ", "")
        data[i][12] = str(data[i][12]).replace(" ", "")
    # print("Les données d'entrée ligne 189 \n", data)
    if not data:
        messagebox.showinfo("Données Manquante", "Il n'y a pas de données exploitables. "
                                                 "Veuillez vérifier le fichiers d'entrée")
        exit()
    else:
        pass
    print("Les données à traitées ont " + str(len(data)) + " lignes")
    # Conversion du champ date[j][3] en string
    for j in range(nb_ligne):
        if isinstance(data[j][3], str):
            print("Ligne 179:", type(data[j][3]))
            pass
        else:
            data[j][3] = data[j][3].strftime('%Y-%m-%d')
            print("Ligne 183:", type(data[j][3]))

    print("les données d'entrée ligne 373 \n", data)
    # exit()
    #########################################

    # Saisie du nom utilisateur et mot de passe
    login = EnterTable4.get()
    mot_de_passe = EnterTable5.get()

    # Saisie de numéro de dossier :
    # numeroDossier = EnterTable6.get()

    # Saisie de la référence de jugement :
    # reference_de_jugement = EnterTable10.get()

    wd_options = Options()
    wd_options.headless = headless
    wd_options.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'
    wd_options.set_preference('detach', True)
    wd = webdriver.Firefox(options=wd_options)
    # wd = webdriver.Firefox(executable_path=GeckoDriverManager().install(), options=wd_options)
    # TODO Passer au service object
    wd.get(
        'https://portailmetierpriv.ira.appli.impots/cas/login?service=http%3A%2F%2Fmedoc.ia.dgfip%3A8141%2Fmedocweb'
        '%2Fcas%2Fvalidation')  # adresse MEDOC DGE

    # wd.get(
    #     'http://medoc.ia.dgfip:8121/medocweb/presentation/md2oagt/ouverturesessionagent/ecran'
    #     '/ecOuvertureSessionAgent.jsf')  # adresse MEDOC Classic
    # Saisir utilisateur
    time.sleep(delay)
    script = f'''identifant = document.getElementById('identifiant'); identifiant.setAttribute('type','hidden'); 
    identifiant.setAttribute('value',"youssef.atigui"); '''
    wd.execute_script(script)

    # Saisie mot de pass
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
    # Saisir service
    wd.find_element(By.ID, 'nomServiceChoisi').send_keys('0070100')  # FRP MEDOC DGE
    # wd.find_element(By.ID, 'nomServiceChoisi').send_keys('6200100')
    time.sleep(delay)
    wd.find_element(By.ID, 'nomServiceChoisi').send_keys(Keys.TAB)

    # Saisir habilitation
    try:
        time.sleep(delay)
        wd.find_element(By.ID, 'habilitation').send_keys('1')
        time.sleep(delay)
        wd.find_element(By.ID, 'habilitation').send_keys(Keys.ENTER)
    except TimeoutException:
        progressbar_label.destroy()
        WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
        messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
        messagebox.showinfo("Service Interrompu !", messages)
        wd.close()

    time.sleep(delay)
    # wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys('3')
    # wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys(Keys.ENTER)
    # time.sleep(delay)

    # Boucle sur le fichier selon le nombre de lignes indiquées
    # for j in range(nb_ligne):
    j = 0
    while True:
        progressbar_label.destroy()
        print("N° de ligne à la ligne 510: ", j)
        num_of_secs = 60
        m, s = divmod(num_of_secs * (nb_ligne + 1), 60)
        min_sec_format = '{:02d}:{:02d}'.format(m, s)
        progressbar_label = Label(tab6,
                                  text=f"Le travail est en cours: {pb['value']:.2f}%  ~  "
                                       f"il reste environ {min_sec_format}")
        progressbar_label.place(x=250, y=label_y)
        tab6.update()
        while True:
            error_messages = ""
            # Création d'un Redevable
            # Arriver à la transactionv 3-17
            try:
                WebDriverWait(wd, 40).until(
                    EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi')))
                wd.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('3')
                wd.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys(Keys.ENTER)
                WebDriverWait(wd, 40).until(
                    EC.presence_of_element_located(
                        (By.ID, 'bmenuxtableMenus:16:outputBmenuxBrmenx04LibelleLigneProposee')))
                wd.find_element(By.ID, 'bmenuxtableMenus:16:outputBmenuxBrmenx04LibelleLigneProposee').click()
            except TimeoutException:
                progressbar_label.destroy()
                messagebox.showinfo("Service Interrompu !",
                                    "La transaction création des oppositions ne semblent pas être "
                                    "disponible. Veuillez tester manuellement avant de redémarrer "
                                    "l'automate.")
                WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                wd.close()
            print("le N° de ligne est à la ligne 547 :", j)
            print("numéro de dossier : ", data[j][0])
            try:
                WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))
                wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[j][0])
                wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(Keys.ENTER)
                print("pas initial")
                break
            except TimeoutException:
                print("messages d'erreur: ", error_messages)
                time.sleep(delay)
                time.sleep(delay)
                time.sleep(delay)
                try:
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
                    error_messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
                    print("messages d'erreur: ", error_messages)
                    message_dossier_verrouille = \
                        "DOSSIER DEJA UTILISE PAR UN AUTRE POSTE  - ATTENTE OU ABANDON - ".replace(" ", "")
                    error_messages_is_present = error_messages.replace(" ", "") == message_dossier_verrouille
                    time.sleep(delay)
                    time.sleep(delay)
                    if error_messages_is_present:
                        messages = f"{error_messages} \n Le dossier N°{data[j][0]} est ouvert par un autre agent ou " \
                                   f"verrouillé.\n Vous pouvez relancer le processus. Cette ligne sera exclu et pourra" \
                                   f" être relancer dans 45 minutes"
                        messagebox.showinfo("Dossier verrouillé !", messages)
                        data[j].append('')
                        data[j].append('\U0001F512')
                        time.sleep(delay)
                        j = j + 1
                        WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                        wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    else:
                        break
                    print("le N° de ligne est à la ligne 579 :", j)
                except TimeoutException:
                    pass
        # Saisie du choix Créer
        try:
            time.sleep(delay)
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI')))
            while wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI') == wd.switch_to.active_element:
                wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys('C')
            print("pas 1")
        except TimeoutException:
            progressbar_label.destroy()
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            # print("ligne 477")
            error_messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messages = error_messages
            print(messages)
            time.sleep(delay)
            time.sleep(delay)
            messagebox.showinfo("Service Interrompu !", messages)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie du numéro de dossier créancier
        try:
            time.sleep(delay)
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))
            wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[j][1])
            wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(Keys.TAB)
            print("pas 2")
            print("le N° de ligne est à la ligne 605 :", j)  # print(data[i][1])
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
            error_messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
            messages = error_messages + "\nLa qualité de la connexion ne permet pas un bon fonctionnement de " \
                                        "l'automate. Veuillez essayer ultérieurement ! "
            messagebox.showinfo("Service Interrompu !", messages)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie de la suite
        try:
            time.sleep(delay)
            time.sleep(delay)
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33gsuitYa33G002ReponseSuite')))
            wd.find_element(By.ID, 'inputB33gsuitYa33G002ReponseSuite').send_keys('S')
            # wd.find_element(By.ID, 'inputB33gsuitYa33G002ReponseSuite').send_keys(Keys.TAB)
            print("pas 3")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # SAISIE DES REFERENCES DE L'OPPOSITION
        # Transport de créance
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance').send_keys('N')
            wd.find_element(By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance').send_keys(Keys.TAB)
            print("pas 4")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie ATD
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GadtAdt')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GadtAdt').send_keys('O')
            wd.find_element(By.ID, 'inputB33ginf2Ya33GadtAdt').send_keys(Keys.TAB)
            print("pas 5")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie du crédit
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GcredCreditIs')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GcredCreditIs').send_keys('N')
            wd.find_element(By.ID, 'inputB33ginf2Ya33GcredCreditIs').send_keys(Keys.TAB)
            print("pas 6")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie Empêchement
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GempEmpechement')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GempEmpechement').send_keys('N')
            wd.find_element(By.ID, 'inputB33ginf2Ya33GempEmpechement').send_keys(Keys.TAB)
            print("pas 7")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie Montant
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GmtMontant')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GmtMontant').send_keys(data[j][2])
            wd.find_element(By.ID, 'inputB33ginf2Ya33GmtMontant').send_keys(Keys.TAB)
            print("pas 8")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie de la Date d'Effet
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
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour').send_keys(date_d_effet.day)
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour').send_keys(Keys.TAB)
            print("pas 9")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie du Mois d'Effet
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois').send_keys(date_d_effet.month)
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois').send_keys(Keys.TAB)
            print("pas 10")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie de l'Année d'Effet
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee').send_keys(date_d_effet.year)
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee').send_keys(Keys.TAB)
            print("pas 11")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie de la référence de jugement
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite')))
            wd.find_element(By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite').send_keys(data[j][4])
            wd.find_element(By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite').send_keys(Keys.TAB)
            print("pas 12")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie de la date d'exécution de jugement
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
            print("pas 13")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie de la date de renouvellement
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
            print("pas 14")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Validation de la non saisie des dates
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON')))
            wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').send_keys('O')
            wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee').send_keys(Keys.TAB)
            print("pas 15")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Validation de la suite
        try:
            time.sleep(delay)
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec')))
            wd.find_element(By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec').send_keys('S')
            wd.find_element(By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec').send_keys(Keys.TAB)
            print("pas 16")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Validation de la saisie de l'opposition
        try:
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33gvlcrYa33GvalcValidationCreation')))
            # time.sleep(delay)
            if wd.find_element(By.ID, 'inputB33gvlcrYa33GvalcValidationCreation').is_displayed():
                wd.find_element(By.ID, 'inputB33gvlcrYa33GvalcValidationCreation').send_keys('O')
                # while wd.find_element(By.ID, 'inputB33gvlcrYa33GvalcValidationCreation').is_displayed():
                #     wd.find_element(By.ID, 'inputB33gvlcrYa33GvalcValidationCreation').send_keys('O')
            print("pas 17")
        # except:
        #     progressbar_label.destroy()
        #     WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
        #     messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
        #     messagebox.showinfo("Service Interrompu !", messages)
        #     WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        #     wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        #     wd.close()
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Capture numéro d'opération
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'outputB33gnopeYa33GnopeNOpe')))
            numero_ope = wd.find_element(By.ID, 'outputB33gnopeYa33GnopeNOpe').text
            print("pas 18")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Saisie de la fin de saisie
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputB33gnouvYa33GnvopNouvelleOpposition')))
            wd.find_element(By.ID, 'inputB33gnouvYa33GnvopNouvelleOpposition').send_keys('N')
            wd.find_element(By.ID, 'inputB33gnouvYa33GnvopNouvelleOpposition').send_keys(Keys.TAB)
            print("pas 19")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Validation de la sortie du formulaire
        try:
            time.sleep(delay)
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            print("pas 20")
        except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
            messagebox.showinfo("Service Interrompu !", message_service_interrompu)
            sav.saved_file(filename=chemin_fichier_de_sortie, j=j, data=data, rep=repertoire_de_sortie,
                           columns=columns_sortie,
                           result='M')
            WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
            wd.find_element(By.ID, 'barre_outils:touche_f2').click()
            wd.quit()
            exit()

        # Marquage tâche faîte dans le fichier
        data[j][13] = numero_ope
        data[j][14] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        data[j][15] = ''
        print("inscription des données dans la liste ligne 789", data)
        print("le N° de ligne est  à la ligne 790:", j)

        # Incrementation ProgressBar

        pb['value'] += num_of_secs / nb_ligne
        progressbar_label.destroy()
        tab6.update()
        progressbar_label = Label(
            tab6, text=f"Le travail est en cours : {pb['value']:.2f}% il reste environ {min_sec_format}")
        progressbar_label.place(x=250, y=label_y)
        pb.update()
        tab6.update()
        print("le N° de ligne est  à la ligne 950:", j)
        print("les nouvelles data : \n", data)
        print("old_data : \n", sortie)
        print("data sans les entêtes (ligne 978)", data)
        numero_affaire = data[j][5]
        sortie = list(filter(lambda x: x[5] != numero_affaire, sortie))
        sortie.append(data[j])
        if sortie[0] != columns_sortie:
            sortie.insert(0, columns_sortie)
        print("listData : \n", sortie)
        save_data(chemin_fichier_de_sortie, sortie)
        if j < nb_ligne - 1:
            j += 1
        else:
            break

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


# Procédure de récupération du numéro d'affaire du RCTVA à imputer
def get_num_affaire(headless=None):
    # Mise en place du format local français
    locale.setlocale(locale.LC_TIME, "fr-FR")
    # Saisie du nom utilisateur et mot de passe
    # login = EnterTable4.get()
    # mot_de_passe = EnterTable5.get()
    login = "meddb-jean-francois.consultant"
    mot_de_passe = "Dagobert01"

    telecharger_CTVA = Telecharger_fichier()

    # Saisie de numéro de dossier :
    # numeroDossier = EnterTable6.get()

    # Saisie de la référence de jugement :
    # reference_de_jugement = EnterTable10.get()
    start = date(2023,datetime.now().month,1)
    end = date(datetime.now().year, datetime.now().month, calendar.monthrange(datetime.now().year, datetime.now().month)[1])
    periode = calendar.monthrange(datetime.now().year, datetime.now().month)[1]
    print(start)
    daterange = []
    for day in range(periode):
        jour = (start + timedelta(days=day))
        if jour.weekday() in [0,1,2,3,4]:
            daterange.append(jour)
    print("liste des jour ouvré du mois",daterange)
    maintenant = datetime.now().date()
    # maintenant = date(2023,9,1)
    indice = daterange.index(maintenant)
    liste_jour_a_telecharger = []
    # nombre de jours entre le jour courant et le 1er jour ouvré du mois courant
    print("la position du jour courant", indice)
    if indice <= 5:
        k = 0
    else:
        k = indice - 4
    for n in range(k,indice):
        print(n)
        jour_a_telecharger = daterange[n]
        liste_jour_a_telecharger.append(jour_a_telecharger)
    liste_jour_a_telecharger.append(maintenant)
    # print(liste_jour_a_telecharger)

    # exit()

    # Téléchargement des données de 5 jours précédent du mois en cours
    # telecharger_CTVA.telecharger(headless,liste_jour_a_telecharger,delay)

    # Vérification de l'existance du repertoire de téléchargement
    telecharge_rep = os.path.expanduser('~')+"\\Downloads"
    if os.path.exists(telecharge_rep):
        # print("MDA" + datetime.now().strftime('_%d_%m_%Y'))
        schema = "MDA"+datetime.now().strftime('_%d_%m_%Y')
        list_fichier_zip = [fichier_zip for fichier_zip in glob.glob(telecharge_rep+"\\*.zip")]
    # # récupération des fichiers zip des 5 jours précedent
        list_TVA_zip = [s for s in list_fichier_zip if schema in s]
        print(list_TVA_zip)
        print(len(list_TVA_zip))
    # # création d'un repertoire d'archive pour les fichiers de crédit de tva
        rep_fichier_tva = os.getcwd()+"\\credit_tva_"+datetime.now().strftime('%d_%m_%Y')
        if not os.path.exists(rep_fichier_tva):
            os.makedirs(rep_fichier_tva)
            print("Repertoire créer")
        else:
            print("Le repertoire existe déjà !")
    # # Ouverture du dernier fichier zip du jour et sauvegarde dans le repertoire
            if len(list_TVA_zip) != 0:
                for zippe in list_TVA_zip:
                    with ZipFile(zippe, 'r') as zip:
                        # afficher tout le contenu du fichier zip
                        zip.printdir()

                        # extraire tous les fichiers
                        print('Extraction...')
                        zip.extractall(rep_fichier_tva)
                        print('Extraction terminé!')
    # # création de l'objet fichier pdf
    # # récupération de la liste des pdfs
    list_fichier_credit_tva = [fichier_credit_tva for fichier_credit_tva in glob.glob(rep_fichier_tva + "\\*.pdf")]
    pdfCreditTvaObj = []
    reader = []
    pdfFileObjList = []
    pageObj = []
    for i in range(len(list_fichier_credit_tva)):
        pdfCreditTvaObj.append(open(list_fichier_credit_tva[i],'rb'))
        reader.append(PyPDF2.PdfReader(pdfCreditTvaObj[i]))
        print("nombres de pages",len(reader[i].pages))
        pdfFileObjList.append(open(list_fichier_credit_tva[i], 'rb'))
    # Création d'une liste d'objet page
        pageObj.append(reader[i].pages[1])

    for j in range(len(pageObj)):
    #  Vérification du code 2
        if "CODIFICATION" and "2  * LE MONTANT A ETE" in pageObj[j].extract_text():
            texte=pageObj[j].extract_text()
            print(texte)
            index_001 = pageObj[j].extract_text().index("001",0,len(pageObj[j].extract_text()))
            index_frp1= pageObj[j].extract_text().index("652271",0,len(pageObj[j].extract_text()))
            index_2023205135 = pageObj[j].extract_text().index("2023205135",0,len(pageObj[j].extract_text()))
            index_002 = pageObj[j].extract_text().index("002", 0, len(pageObj[j].extract_text()))
            index_frp2 = pageObj[j].extract_text().index("632232", 0, len(pageObj[j].extract_text()))
            index_2023205136 = pageObj[j].extract_text().index("2023205136", 0, len(pageObj[j].extract_text()))
            print("l'index de la 1er ligne",index_001)
            print("index frp1", index_frp1)
            print("index du N° affaire 1",index_2023205135)
            print("l'index de la 2eme ligne",index_002)
            print("index frp2", index_frp2)
            print("index du N° affaire 2",index_2023205136)
            numero_affaire1 = texte[index_2023205135:index_2023205135+10]
            numero_affaire2 = texte[index_2023205136:index_2023205136 + 10]
            print("le numéro d'affaire 1 est:", numero_affaire1)
            print("le numéro d'affaire 2 est:", numero_affaire2)
    # Détermination du nombre de ligne à traité dans le fichier
            texte_nombre = texte.index("NOMBRE DE DEMANDES EN AFFAIRE")
            l = texte[texte_nombre:texte_nombre+53]
            x = re.findall("[0-9]+",l)
            print("le nombre de tour est de :", x[0])
            numero_affaire: list[str] = []
            for m in range(int(x[0])):
                numero_affaire.append(texte[2562+m*552:+2562+m*552+10])
                print(m)
            print("la liste des numero d'affaires est:", numero_affaire)
            message = "Les N° d'affaires en code 2 pour la période allant du "\
                          +liste_jour_a_telecharger[0].strftime("%A %d %B") + " jusqu'à ce jour, sont :\n\u2022 "+'\n'  u'\u2022 '.join(numero_affaire)
            text_box = Text(
                tab1,
                height=9,
                width=70,
                wrap='word',
                font=('Arial', 13)
                )
            text_box.place(x=250, y=120)
            text_box.insert('end', message)
            text_box.config(state='disabled')

    # Analyse des document charger


    print("fin du programme")


# Procédure pour la vérification du fichier
def open_file():
    global File_path
    # global l1
    global nb_ligne1
    df = pd.DataFrame()
    source_rep = os.getcwd()
    file = filedialog.askopenfile(mode='r', filetypes=[('Ods Files', '*.ods')],
                                  initialdir='C:\\Users\\Meddb-jean-francoi01\\Documents\\automate_satd\\entrees_SATD')
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
        print("le dataframe des anciennes données : \n", df1)
        print("----------------------------------------------------------------------------")
        nb_ligne1 = df1.shape[0]
        sub_df1 = df1[df1['Dossiers traités'] == '']
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
                "Création d'opposition", "Vous avez déjà effectué les opérations sur ce fichier."
                                         "\n Voulez-vous continuer")
            if not response:
                Interface.destroy()
            else:
                pass

    file.close()
    for i in range(df.shape[0]):
        colonne_13 = "Numéro et date de l'opération de dépense effectuée dans Médoc pour paiement du poste " \
                     "comptable RNF ayant émis la SATD "
        colonne_14 = "Dossiers traités"
        if df.drop(columns=[colonne_13, colonne_14]).loc[i].isnull().any():
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
label1 = Label(tab1, text='Récupération des numéro d\'affaire', font=('Arial', 15), fg='Black', bg='#ffffff',
               relief="sunken")
label1.place(x=350, y=paramx)

tab2 = Frame(tabControl, bg='#E3EBD0')
label2 = Label(tab2, text='Créer des oppositions', font=('Arial', 15), fg='Black', bg='#ffffff', relief="sunken")
label2.place(x=400, y=paramx)
# tabControl.add(tab1, text='Liste des oppositions')
# tabControl.add(tab2, text='Création des oppositions')
tabControl.pack(expand=1, fill="both")
tab6 = Frame(tabControl, bg='#E3EBD0')
tabControl.add(tab6, text='Création d\'opposition')
tabControl.add(tab1, text='Récupération des numéro d\'affaire')
tabControl.pack(expand=1, fill="both")
tab3 = Frame(tabControl, bg='#E3EBD0')
tab4 = Frame(tabControl, bg='#E3EBD0')

# Etablissement de l'image de fermeture
# img = Image.open('close-button.png')
# img_resize = img.resize((30, 30), Image.LANCZOS)
# closeIcon = ImageTk.PhotoImage(img_resize)
# closeButton1 = Button(Interface, image=closeIcon, command=lambda: tabControl.forget(tab3))
# closeButton1.pack(side=LEFT)
# closeButton2 = Button(Interface, image=closeIcon, command=lambda: tabControl.forget(tab4))
# closeButton2.pack(side=LEFT)

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

lexique = "Lexique : \n    ● Le symbole 'X' indique que la ligne a été traitée avec succès.\n    ● Le " \
          "symbole '∅' indique que des données obligatoires sont manquantes sur la ligne en question. \n    ● Le " \
          "symbole '\U0001F512' indique que le dossier de la ligne en question et verrouillé, la ligne pourra " \
          "être retraitée dans un délai de 45 minutes.\n Pour traitées les lignes comportant des anomalies, " \
          "vous avez juste à relancé l’automates une fois les anomalies résolues. "
labelLexique = Label(tab6, text=lexique, relief="sunken", wraplength=500, justify=LEFT)
labelLexique.place(x=250, y=paramy + 235)

# login et mot de passe sur tab1 à tab3
label5 = Label(tab1, text='Identifiant:', relief="sunken")
label5.place(x=250, y=70)
entry4 = Entry(tab1, textvariable=EnterTable4, justify='center', width=30)
entry4.place(x=340, y=70)
label6 = Label(tab1, text='Mot de passe: ', relief="sunken")
label6.place(x=550, y=70)
mot_de_passe = Entry(tab1, textvariable=EnterTable5, show="*", justify='center')
mot_de_passe.place(x=650, y=70)

browser_button = Button(tab1, bg="#82CFD8", text='Récuperer les numéro d\'affaires RCTVA sans visualisation',
                        command=lambda: get_num_affaire(headless=True))
browser_button.place(x=paramx + 240, y=paramy + 250)

recup_num_affaire = Button(tab1, bg="#007FA9", text='Récuperer les numéro d\'affaires RCTVA avec visualisation',
                           command=lambda: get_num_affaire(headless=False))
recup_num_affaire.place(x=paramx + 240, y=paramy + 150)

label5 = Label(tab2, text='Identifiant:', relief="sunken")
label5.place(x=250, y=70)
entry4 = Entry(tab2, textvariable=EnterTable4, justify='center')
entry4.place(x=340, y=70)
label6 = Label(tab2, text='Mot de passe: ', relief="sunken")
label6.place(x=500, y=70)
mot_de_passe = Entry(tab2, textvariable=EnterTable5, show="*", justify='center')
mot_de_passe.place(x=600, y=70)


def toggle_password():
    if mot_de_passe.cget('show') == '':
        time.sleep(10)
        mot_de_passe.config(show='*')
    else:
        mot_de_passe.config(show='')
        button_font.config(overstrike=1)


button_font = font.Font(family='Tahoma', size=12)
show_password_btn = Button(tab1, text='👁', font=button_font, justify=CENTER, command=toggle_password)
show_password_btn.place(x=600 + 70 + 135, y=65)

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
mot_de_passe = Entry(tab6, textvariable=EnterTable5, justify='center')
mot_de_passe.place(x=600, y=70)

button2 = Button(tab6, bg="#CEDDDE", text='Choisir le fichier d\'entrée', command=open_file)
button2.place(x=paramx + 240, y=paramy - 30)
label_path6 = Label(tab6)
label_path6.place(x=paramx + 490, y=paramy - 30)

browser_button = Button(tab6, bg="#82CFD8", text='Créer les Oppositions sans visualisation des transactions',
                        command=lambda: create_opposition(headless=True))
browser_button.place(x=paramx + 240, y=paramy + 100)
creerOpposition = Button(tab6, bg="#007FA9", text='Créer les Oppositions avec visualisation des transactions',
                         command=lambda: create_opposition(headless=False))
creerOpposition.place(x=paramx + 240, y=paramy + 150)

Interface.mainloop()
