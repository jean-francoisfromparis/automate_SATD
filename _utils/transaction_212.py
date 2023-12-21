import os
import sys
import time
from datetime import datetime, timedelta
from tkinter import messagebox, Label

import pandas as pd
from pyexcel_ods import save_data
from selenium.webdriver.common import keys
from selenium.common import TimeoutException, StaleElementReferenceException, \
    ElementNotInteractableException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from _utils.save_file import Saved_file


class Transaction_212:
    def transaction_212(self, headless):
        sav = Saved_file()
        delay = 5
        File_path = 'C:\\Users\\Meddb-jean-francoi01\\Documents\\automate_satd\\donnees_sortie_phase1\\donnees_sortie_phase1_2023-11-24\\donnees_sortie_phase1_2023-11-24.ods'
        source_rep = os.getcwd()
        fichier_de_Sortie = 'donnees_sortie_phaseII' + datetime.now().strftime('_%Y-%m-%d') + '.ods'
        sortie_repertoire = source_rep + '/donnees_sortie_phaseII/sorties_SATD' + datetime.now().strftime('_%Y-%m-%d')
        os.mkdir(sortie_repertoire)
        df = pd.read_excel(File_path)
        columns = df.columns.values.tolist()
        columns_sortie = columns + ["Numéro d'Opération phase II", "Date d'exécution phase II",
                                    "Dossiers traités phase II"]
        wd_options = Options()
        if headless:
            wd_options.add_argument("--headless")
        else:
            pass
        wd_options.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'
        wd_options.set_preference('detach', True)
        wd = webdriver.Firefox(options=wd_options)

        wd.get(
            'https://portailmetierpriv.ira.appli.impots/cas/login?service=http%3A%2F%2Fmedoc.ia.dgfip%3A8141%2Fmedocweb'
            '%2Fcas%2Fvalidation')  # adresse MEDOC DGE
        while wd.title == "Identification":
            # Saisir utilisateur
            time.sleep(delay)
            script = f'''identifant = document.getElementById('identifiant'); identifiant.setAttribute('type','hidden');
            identifiant.setAttribute('value',"youssef.atigui"); '''
            wd.execute_script(script)
            time.sleep(delay)
            # wd.find_element(By.ID, 'identifiant').send_keys(login)
            # print(login)

            # Saisie mot de passe
            time.sleep(delay)
            # wd.find_element(By.ID, 'secret_tmp').send_keys(mot_de_passe)
            # print(mot_de_passe)
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

            time.sleep(delay)
            wd.find_element(By.ID, 'nomServiceChoisi').send_keys(Keys.TAB)

            # Saisir habilitation
            try:
                time.sleep(delay)
                wd.find_element(By.ID, 'habilitation').send_keys('1')
                time.sleep(delay)
                wd.find_element(By.ID, 'habilitation').send_keys(Keys.ENTER)
            except TimeoutException:
                # progressbar_label.destroy()
                WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
                messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
                messagebox.showinfo("Service Interrompu !", messages)
                wd.close()
            data = df.values.tolist()
            nb_ligne = len(data)
            print("Les données", data)
            j = 0
            while True:
                start = time.time()
                num_of_secs = 540
                min, sec = divmod(num_of_secs * 3, 60)
                hour, min = divmod(min, 60)
                heure_demarrage = datetime.now()
                heure_de_fin = heure_demarrage + timedelta(hours=hour, minutes=min)
                heure_de_fin = heure_de_fin.strftime('%H:%M:%S')
                # pb = progressbar(tab2)
                # progressbar_label = Label(tab2,
                #                           text=f"Le travail est en cours: {pb['value']:.2f}%  ~  "
                #                                f"L'opération de création de SATD sera terminé à  {heure_de_fin}")
                # progressbar_label.place(x=250, y=470)
                # tab2.update()

                # Saisie de la transaction 21-2
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 10).until(
                        EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte')))
                    wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys('212')
                    wd.find_element(By.ID, 'inputBmenuxBrmenx07CodeSaisieDirecte').send_keys(Keys.ENTER)
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Création affaire service au code R17 "7055"
                # Saisie de la nature "AFF" pour debit 473-0
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(
                        EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('AFF')
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)
                    self.Deb("pas 1")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                    # Saisie du type de montant
                try:
                    WebDriverWait(wd, 40).until(
                        EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
                    print(data[j][2])
                    self.Deb("pas 2")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                    # Saisie du montant X
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(
                        EC.presence_of_element_located(
                            (By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(
                        int(data[j][2]))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(
                        Keys.ENTER)
                    self.Deb("pas 3")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir une identification
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(
                        EC.presence_of_element_located(
                            (By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep')))
                    wd.find_element(By.ID,
                                    'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(
                        data[j][4])
                    wd.find_element(By.ID,
                                    'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(
                        Keys.ENTER)
                    self.Deb("pas 4")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisie du numéro d'affaire
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(
                        EC.presence_of_element_located((By.ID, 'inputBcaff03Nuaff1NumeroAffaire')))
                    wd.find_element(By.ID, 'inputBcaff03Nuaff1NumeroAffaire').send_keys(data[j][5])
                    print("numèro affaire:", data[j][5])
                    self.Deb("pas 5")
                    time.sleep(delay)
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):

                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Confirmer le libelle de l'affaire
                try:
                    time.sleep(delay)
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(
                        EC.presence_of_element_located((By.ID, 'inputBcaff03LaffAffLibelleAffaire')))
                    wd.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(Keys.ENTER)
                    self.Deb("pas 6")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                    # Saisir le code R27 "7370"
                try:
                    time.sleep(delay)
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi')))
                    wd.find_element(By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi').send_keys(Keys.ENTER)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01RMontantSaisi')))
                    time.sleep(delay)
                    wd.find_element(By.ID, 'inputBcaff01Bcaff01RMontantSaisi').send_keys(Keys.ENTER)
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01SValidationOperateur')))
                    wd.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys('O')
                    wd.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys(Keys.ENTER)
                    time.sleep(delay)
                    # WebDriverWait(wd, 40).until(
                    #     EC.presence_of_element_located((By.ID, 'inputBcaff12Bcaff121ValidationON')))
                    # wd.find_element(By.ID, 'inputBcaff12Bcaff121ValidationON').send_keys('O')
                    # wd.find_element(By.ID, 'inputBcaff12Bcaff121ValidationON').send_keys(Keys.ENTER)
                    self.Deb("pas 7")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir le numéro du compte 477-0
                try:
                    time.sleep(delay)
                    time.sleep(delay)
                    time.sleep(delay)
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(
                        EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI')))
                    print("ok")
                    time.sleep(delay)
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI').send_keys(Keys.ENTER)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx')))
                    time.sleep(delay)
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx').send_keys('477-0')
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx').send_keys(Keys.ENTER)
                    self.Deb("pas 8")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir la nature "AFF" pour crédit 477-0
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(
                        EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('AFF')
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)
                    self.Deb("pas 9")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                    # Saisie du type de montant
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
                    self.Deb("pas 10")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisie du montant X
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(data[j][7])
                    time.sleep(delay)
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(Keys.ENTER)
                    self.Deb("pas 11")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisie la date comptable
                # Capture et réutilisation de la date journée comptable
                try:
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'PDATCPT_dateJourneeComptable')))
                    djc_capture = wd.find_element(By.ID, 'PDATCPT_dateJourneeComptable').text
                    djc = djc_capture.split('/')
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation').send_keys(djc[0])
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation').send_keys(Keys.ENTER)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation').send_keys(djc[1])
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation').send_keys(Keys.ENTER)
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation').send_keys(djc[2])
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation').send_keys(Keys.ENTER)
                    self.Deb("pas 12")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisie du numéro d'affaire
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff03Nuaff1NumeroAffaire')))
                    wd.find_element(By.ID, 'inputBcaff03Nuaff1NumeroAffaire').send_keys(Keys.ENTER)
                    self.Deb("pas 13")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisie le numéro de dossier
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff03Bcaff032RedevServOuRlce')))
                    wd.find_element(By.ID, 'inputBcaff03Bcaff032RedevServOuRlce').send_keys('REDEV')
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff03Bcaff036Car2A7NuordNumDos')))
                    wd.find_element(By.ID, 'inputBcaff03Bcaff036Car2A7NuordNumDos').send_keys(data[j][0])
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(
                        EC.presence_of_element_located((By.ID, 'inputBcaff03Bcaff038Cplnum')))
                    wd.find_element(By.ID, 'inputBcaff03Bcaff038Cplnum').send_keys('0')
                    wd.find_element(By.ID, 'inputBcaff03Bcaff038Cplnum').send_keys(Keys.ENTER)
                    self.Deb("pas 14")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisie du libellé
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(
                        EC.presence_of_element_located((By.ID, 'inputBcaff03LaffAffLibelleAffaire')))
                    wd.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(data[j][4])
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(
                        EC.presence_of_element_located((By.ID, 'inputBcaff03LaffAffLibelleAffaire')))
                    wd.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(Keys.ENTER)
                    self.Deb("pas 15")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir le code R27 "7055"
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi')))
                    wd.find_element(By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi').send_keys(Keys.ENTER)
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff01Cr17R27CodeR17OuR27')))
                    wd.find_element(By.ID, 'inputBcaff01Cr17R27CodeR17OuR27').send_keys('7055')
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01RMontantSaisi')))
                    wd.find_element(By.ID, 'inputBcaff01Bcaff01RMontantSaisi').send_keys(data[j][7])
                    wd.find_element(By.ID, 'inputBcaff01Bcaff01RMontantSaisi').send_keys(Keys.ENTER)
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01SValidationOperateur')))
                    wd.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys('O')
                    self.Deb("pas 16")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Validation de la transaction
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvim01Ycvim013ReponseOperateur')))
                    wd.find_element(By.ID, 'inputBcvim01Ycvim013ReponseOperateur').send_keys('O')
                    self.Deb("pas 17")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Création d'une liste temporaire avec numéro d'ordre de dépenses, le numéro de l'affaire créée et le numéro
                # de l'opération Le numéro de l'opération est divisé sur deux cellules dans MEDOC Cette liste sera finalement
                # collée comme ligne dans le fichier des donnees de sortie
                liste_temporaire_data = [str(data[j][0]),
                                         str(data[j][7])]  # FRP indice #0 dans liste_temporaire_data
                # #Montant indice #1 dans liste_temporaire_data

                # Numéro de l'ordre de dépense 1
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'outputBcvcs04Ycvcs022NoDepense')))
                    # Numéro de l'ordre de dépense indice #2 dans liste_temporaire_data
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2')))
                    liste_temporaire_data.append(wd.find_element(By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2').text)
                    self.Deb("pas 18")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Pour afficher la suite
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvcs04Ycvcs028Reponse')))
                    wd.find_element(By.ID, 'inputBcvcs04Ycvcs028Reponse').send_keys(Keys.ENTER)
                    self.Deb("pas 19")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Numéro de l'affaire créée
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'outputBcvcs04Ycvcs021NumAffaireCreee')))
                    # Numéro de l'affaire créée indice #3 dans liste_temporaire_data
                    numero_affaire_creee = wd.find_element(By.ID, 'outputBcvcs04Nuaff1NumeroAffaire').text
                    liste_temporaire_data.append(numero_affaire_creee)
                    self.Deb("pas 20")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                if liste_temporaire_data[3] != numero_affaire_creee or liste_temporaire_data[3] == '':
                    time.sleep(delay)
                    numero_affaire_creee_v = wd.find_element(By.ID, 'outputBcvcs04Nuaff1NumeroAffaire').text
                    liste_temporaire_data[3] = numero_affaire_creee_v
                else:
                    pass

                # Pour afficher la suite
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvcs04Ycvcs028Reponse')))
                    wd.find_element(By.ID, 'inputBcvcs04Ycvcs028Reponse').send_keys(Keys.ENTER)
                    self.Deb("pas 21")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Numero de l'opération 1
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'labelBcvcs032')))
                    # Numero de l'opération indice #4 dans liste_temporaire_data
                    liste_temporaire_data.append(wd.find_element(By.ID, 'outputBcvcs03Nuopet1ErCarNuopeF2').text +
                                                 wd.find_element(By.ID, 'outputBcvcs03Nuopes5DerniersCarNuope').text)
                    self.Deb("pas 22")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Pour afficher la suite
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvcs03Ycvcs014DemandeSuite')))
                    wd.find_element(By.ID, 'inputBcvcs03Ycvcs014DemandeSuite').send_keys('S')
                    self.Deb("pas 23")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Fin de la transaction 21-2 et retour à la page d'accueil
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'barre_outils:image_f2')))
                    wd.find_element(By.ID, 'barre_outils:image_f2').click()
                    self.Deb("pas 24")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir la transaction 21-2
                ## Saisie de la transaction 21-2
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 10).until(EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi')))
                    wd.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')
                    time.sleep(delay)
                    wd.find_element(By.ID, 'inputBmenuxBrmenx062ECaractere').send_keys('1')
                    time.sleep(delay)
                    WebDriverWait(wd, 10).until(EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi')))
                    wd.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')
                    self.Deb("pas 25")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # progressbar_label.destroy()
                    # WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
                    # messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
                    # messages = "Une erreur inattendu"
                    # messagebox.showinfo("Service Interrompu !", messages)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Création affaire service au code R27 "8755"
                # Saisir la nature "AFF" pour debit 473-0
                # Saisir la nature "AFF" pour crédit 477-0
                try:
                    time.sleep(delay)
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('AFF')
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)
                    self.Deb("pas 26")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire, columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisie du type de montant
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
                    self.Deb("pas 27")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisie du montant X
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(data[j][7])
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(Keys.ENTER)
                    self.Deb("pas 28")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir une identification
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep')))
                    wd.find_element(By.ID,'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(data[j][8])
                    wd.find_element(By.ID,'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(Keys.ENTER)
                    self.Deb("pas 29")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir le numéro d'affaire créée précédemment
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff03Nuaff1NumeroAffaire')))
                    wd.find_element(By.ID, 'inputBcaff03Nuaff1NumeroAffaire').send_keys(liste_temporaire_data[3])
                    self.Deb("pas 30")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Confirmer le libelle de l'affaire
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff03LaffAffLibelleAffaire')))
                    wd.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(Keys.ENTER)
                    self.Deb("pas 31")  # print("pas 31 - ligne 1276")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir le code R27 "8755"
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi')))
                    wd.find_element(By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi').send_keys(Keys.ENTER)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01RMontantSaisi')))
                    wd.find_element(By.ID, 'inputBcaff01Bcaff01RMontantSaisi').send_keys(Keys.ENTER)
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01SValidationOperateur')))
                    wd.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys('O')
                    self.Deb("pas 32")  # print("pas 32 - ligne 1298")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Répondre à la question "Soldez-vous l'affaire ?"
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff12Bcaff121ValidationON')))
                    wd.find_element(By.ID, 'inputBcaff12Bcaff121ValidationON').send_keys('O')
                    self.Deb("pas 33")  # print("pas 33 - ligne 1314")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Valider CREDIT
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI').send_keys(Keys.ENTER)
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx').send_keys('512-96')
                    self.Deb("pas 34")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir la nature "OVIRT"
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('OVIRT')
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)
                    self.Deb("pas 35")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir ENTREE pour type de montant
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
                    self.Deb("pas 36")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisie du montant X
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(data[j][7])
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(Keys.ENTER)
                    self.Deb("pas 27 - ligne 862")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisie du codique du service bénéficiaire
                try:
                    time.sleep(delay)
                    print("ok")
                    # WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.ID, 'inputBnN4f3001Bn4F300101ZoneCodiqueService')))
                    # wd.find_element(By.ID, 'inputBn4f3001Bn4F300101ZoneCodiqueService').send_keys(data[j][9])
                    # script = "document.getElementById('inputBnN4f3001Bn4F300101ZoneCodiqueService').setAttribute('value', data[j][9])"
                    # wd.execute_script(script)
                    WebDriverWait(wd, 80).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[6]/form[4]/div/div[3]/table[7]/tbody/tr[2]/td[6]/input')))
                    time.sleep(delay)
                    wd.find_element(By.XPATH,'/html/body/div[2]/div[6]/form[4]/div/div[3]/table[7]/tbody/tr[2]/td[6]/input').send_keys(
                        str(data[0][9]).rjust(2 + len(str(data[0][9])), '0'))
                    print("data 9:", str(data[0][9]).rjust(2 + len(str(data[0][9])), '0'))
                    self.Deb("pas 38")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Appuyer sur Entrer pour continuer
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBn4f3001Bn4F300116ZoneAcquisitionLibre')))
                    wd.find_element(By.ID, 'inputBn4f3001Bn4F300116ZoneAcquisitionLibre').send_keys(Keys.ENTER)
                    self.Deb("pas 39")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Validation de la transaction
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvim01Ycvim013ReponseOperateur')))
                    wd.find_element(By.ID, 'inputBcvim01Ycvim013ReponseOperateur').send_keys('O')
                    self.Deb("pas 40")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Numéro de l'ordre de dépense 2
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2')))
                    # Numero de l'ordre de depense indice #5 dans liste_temporaire_data
                    liste_temporaire_data.append(wd.find_element(By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2').text)
                    self.Deb("pas 41")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Pour afficher la suite
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvcs04Ycvcs028Reponse')))
                    wd.find_element(By.ID, 'inputBcvcs04Ycvcs028Reponse').send_keys(Keys.ENTER)
                    self.Deb("pas 42")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Numéro de l'opération 2
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'labelBcvcs032')))
                    # Numero de l'opération indice #6 dans liste_temporaire_data
                    liste_temporaire_data.append(wd.find_element(By.ID, 'outputBcvcs03Nuopet1ErCarNuopeF2').text +
                                                  wd.find_element(By.ID, 'outputBcvcs03Nuopes5DerniersCarNuope').text)
                    self.Deb("pas 43")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Pour afficher la suite
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvcs03Ycvcs014DemandeSuite')))
                    wd.find_element(By.ID, 'inputBcvcs03Ycvcs014DemandeSuite').send_keys('S')
                    self.Deb("pas 44")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Fin de la transaction 21-2 et retour à la page d'accueil
                time.sleep(delay)
                WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'barre_outils:image_f2')))
                wd.find_element(By.ID, 'barre_outils:image_f2').click()
                self.Deb("pas 45")

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
                    self.Deb("pas 46")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir le numéro d'affaire à partir des données d'entrées
                try:
                    time.sleep(delay)
                    time.sleep(delay)
                    time.sleep(delay)
                    set_numero_affaire = f'''document.getElementById('inputBrsdo03Nuaff1NumeroAffaire').setAttribute('value','{data[j][5]}'); '''
                    wd.execute_script(set_numero_affaire)
                    time.sleep(delay)
                    wd.find_element(By.XPATH,'/html/body/div[2]/div[6]/form[4]/div/div[3]/table/tbody/tr[2]/td[4]/input').send_keys(Keys.ENTER)
                    time.sleep(delay)
                    time.sleep(delay)
                    time.sleep(delay)
                    self.Deb("pas 47")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir le type de l'affaire "64"
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBrsdo03NasdoNatureSousDossier')))
                    wd.find_element(By.ID, 'inputBrsdo03NasdoNatureSousDossier').send_keys('64')
                    self.Deb("pas 48")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Récuperer le nouveau solde de l'affaire au code 1760 et enregistrer le sous indice #7 dans liste_temporaire_data
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'outputBraff01Yraff01YSoldeArticle')))
                    liste_temporaire_data.append(wd.find_element(By.ID, 'outputBraff01Yraff01YSoldeArticle').text)
                    self.Deb("pas 49")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Récupérer le nom de l'entreprise à rembourser et enregistrer le dans une liste temporaire
                liste_tempo_nom_entreprise = [str(data[j][0])]
                liste_tempo_nom_entreprise.append(str(data[j][0]))
                WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'outputBrtit04NomprfNomProfession')))
                liste_tempo_nom_entreprise.append(wd.find_element(By.ID, 'outputBrtit04NomprfNomProfession').text + "/SOLDE RCTVA")
                self.Deb("pas 50")

                # Pour afficher la suite
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'labelBrval18BarreEspace0')))
                    wd.find_element(By.ID, 'inputYrval18wAcquisitionEspace').send_keys(Keys.ENTER)
                    self.Deb("pas 51")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Pour afficher la suite encore une fois en cas de besoin
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputYrval18wAcquisitionEspace')))
                    wd.find_element(By.ID, 'inputYrval18wAcquisitionEspace').send_keys(Keys.ENTER)
                    self.Deb("pas 52")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Fin de la transaction 3-8-2 et retour à la page d'accueil
                time.sleep(delay)
                WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'barre_outils:image_f2')))
                wd.find_element(By.ID, 'barre_outils:image_f2').click()
                self.Deb("pas 53")

                # Saisir la transaction 21-2
                # Remboursement du solde à la société débitrice
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi')))
                    wd.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx062ECaractere')))
                    wd.find_element(By.ID, 'inputBmenuxBrmenx062ECaractere').send_keys('1')
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi')))
                    wd.find_element(By.ID, 'inputBmenuxBrmenx051ErCaractereSaisi').send_keys('2')
                    self.Deb("pas 54")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # progressbar_label.destroy()
                    WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
                    messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
                    messagebox.showinfo("Service Interrompu !", messages)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir la nature "AFF" pour debit 473-0
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('AFF')
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)
                    self.Deb("pas 55")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisie du type de montant
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 10).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
                    self.Deb("pas 56")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

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
                wd.find_element(By.XPATH,'/html/body/div[2]/div[6]/form[4]/div/div[3]/table[1]/tbody/tr[3]/td[18]/input').send_keys(Keys.ENTER)
                self.Deb("pas 57")

                # Saisie de l'identification
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep')))
                    wd.find_element(By.ID,'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(liste_tempo_nom_entreprise[1])
                    time.sleep(delay)
                    wd.find_element(By.ID,'repeatBcimp01:0:inputBciddepPc8IdentIndentificationBeneficiaireDep').send_keys(Keys.ENTER)
                    self.Deb("pas 58")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisie du numéro d'affaire créée précédemment
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff03Nuaff1NumeroAffaire')))
                    wd.find_element(By.ID, 'inputBcaff03Nuaff1NumeroAffaire').send_keys(data[j][5])
                    self.Deb("pas 59")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Confirmer le libelle de l'affaire
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff03LaffAffLibelleAffaire')))
                    wd.find_element(By.ID, 'inputBcaff03LaffAffLibelleAffaire').send_keys(Keys.ENTER)
                    self.Deb("pas 60")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir le code R27 "7370"
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi')))
                    wd.find_element(By.ID, 'inputBcaff01Bcaff01PNumeroLigneSaisi').send_keys(Keys.ENTER)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01RMontantSaisi')))
                    wd.find_element(By.ID, 'inputBcaff01Bcaff01RMontantSaisi').send_keys(Keys.ENTER)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcaff01Bcaff01SValidationOperateur')))
                    wd.find_element(By.ID, 'inputBcaff01Bcaff01SValidationOperateur').send_keys('O')
                    self.Deb("pas 61")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Répondre à la question "Soldez-vous l'affaire ?"
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.ID, 'inputBcaff12Bcaff121ValidationON')))
                    wd.find_element(By.ID, 'inputBcaff12Bcaff121ValidationON').send_keys('O')
                    self.Deb("pas 60")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Valider CREDIT
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp011ActionCSOuI').send_keys(Keys.ENTER)
                    self.Deb("pas 61")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir le numéro du compte 512-96
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01CcompteNumeroCompteXxxXx').send_keys('512-96')
                    self.Deb("pas 62")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisie de la nature "VIRT"
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys('VIRT')
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp013NatureSaisie').send_keys(Keys.ENTER)
                    self.Deb("pas 63")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisie du type de montant
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01YcimpdevCEstDeviseEOuF').send_keys(Keys.ENTER)
                    self.Deb("pas 64")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisie du montant X
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(liste_temporaire_data[7])
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp018MontantImputation').send_keys(Keys.ENTER)
                    self.Deb("pas 67")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisie de la date du jour comptable
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation').send_keys(djc[0])
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016JourDateImputation').send_keys(Keys.ENTER)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation').send_keys(djc[1])
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016MoisDateImputation').send_keys(Keys.ENTER)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation')))
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation').send_keys(djc[2])
                    wd.find_element(By.ID, 'repeatBcimp01:0:inputBcimp01Ycimp016AnneeDateImputation').send_keys(Keys.ENTER)
                    self.Deb("pas 68")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisie du numéro de dossier
                try:
                    time.sleep(delay)
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))
                    wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[j][0])
                    if wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').is_displayed:
                        wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[j][0])
                    self.Deb("pas 69")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Continuer en cas d'existence de ce message : ATTENTION - OPPOSITION POUR CE DOSSIER
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_all_elements_located((By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON')))
                    wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').send_keys('O')
                    self.Deb("pas 70")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

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
                    WebDriverWait(wd, 40).until(EC.presence_of_all_elements_located((By.ID, 'inputBibanremYaribmess1LibelleMessage')))
                    wd.find_element(By.ID, 'inputBibanremYaribchoixSaisieChoix').send_keys(data[j][11])
                    self.Deb("pas 71")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Libelle du virement emis
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_all_elements_located((By.ID, 'repeatBcimp01:1:inputBcrib01IbanlibLibelleVirementEmis')))
                    wd.find_element(By.ID, 'repeatBcimp01:1:inputBcrib01IbanlibLibelleVirementEmis').send_keys(str(data[j][12]) + "/ RCTVA")
                    wd.find_element(By.ID, 'repeatBcimp01:1:inputBcrib01IbanlibLibelleVirementEmis').send_keys(Keys.ENTER)
                    self.Deb("pas 72")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Répondre à la question "Voulez-vous valider ?"
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvim01Ycvim013ReponseOperateur')))
                    wd.find_element(By.ID, 'inputBcvim01Ycvim013ReponseOperateur').send_keys('O')
                    self.Deb("pas 73")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Visualisation du Numero de l'ordre de dépense
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2')))
                    # Numero de l'ordre de dépense indice #8 dans liste_temporaire_data
                    liste_temporaire_data.append(
                        wd.find_element(By.ID, 'outputBcvcs04NudepNumeroPieceDepenseCfF2').text)
                    self.Deb("pas 73")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Pour afficher la suite
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvcs04Ycvcs028Reponse')))
                    wd.find_element(By.ID, 'inputBcvcs04Ycvcs028Reponse').send_keys(Keys.ENTER)
                    self.Deb("pas 74")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Numero de l'opération
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'labelBcvcs032')))
                    # Numero de l'opération indice #9 dans liste_temporaire_data
                    liste_temporaire_data.append( wd.find_element(By.ID, 'outputBcvcs03Nuopet1ErCarNuopeF2').text + wd.find_element(By.ID, 'outputBcvcs03Nuopes5DerniersCarNuope').text)
                    self.Deb("pas 75")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Pour afficher la suite
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputBcvcs03Ycvcs014DemandeSuite')))
                    wd.find_element(By.ID, 'inputBcvcs03Ycvcs014DemandeSuite').send_keys('S')
                    self.Deb("pas 76")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

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
                    self.Deb("pas 77")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Création affaire service au code R17 "7055"
                # Saisir le numéro du dossier
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputYrdos211NumeroDeDossier')))
                    wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[j][0])
                    self.Deb("pas 78")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir M pour mise à jour
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'labelB33gmenu0')))
                    wd.find_element(By.ID, 'inputB33gmenuYa33Gch1ChoixCMAI').send_keys("M")
                    self.Deb("pas 79")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir le numéro FRP de l'opposant
                try:
                    time.sleep(delay)
                    wd.find_element(By.ID, 'inputYrdos211NumeroDeDossier').send_keys(data[j][1])
                    self.Deb("pas 80")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

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
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf1Ya33GcpaComplementAdresse')))
                    wd.find_element(By.ID, 'inputB33ginf1Ya33GcpaComplementAdresse').send_keys(Keys.ENTER)
                    # Adresse CODE POSTAL
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf1Ya33GcpoCodePostal')))
                    wd.find_element(By.ID, 'inputB33ginf1Ya33GcpoCodePostal').send_keys(Keys.ENTER)
                    # Adresse BUREAU
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf1Ya33GbudiBureau')))
                    wd.find_element(By.ID, 'inputB33ginf1Ya33GbudiBureau').send_keys(Keys.ENTER)
                    self.Deb("pas 81")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir S pour la validation
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33gsuvlYa33G009Reponse')))
                    wd.find_element(By.ID, 'inputB33gsuvlYa33G009Reponse').send_keys("S")
                    self.Deb("pas 82")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Taper sur la touche «Entrée» jusqu’à la case de saisie
                # Transport de creance O/N
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance')))
                    wd.find_element(By.ID, 'inputB33ginf2Ya33GtrcrTransportCreance').send_keys(Keys.ENTER)

                    # ATD, Saisies O/N
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GadtAdt')))
                    wd.find_element(By.ID, 'inputB33ginf2Ya33GadtAdt').send_keys(Keys.ENTER)

                    # Crédit O/N
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GcredCreditIs')))
                    wd.find_element(By.ID, 'inputB33ginf2Ya33GcredCreditIs').send_keys(Keys.ENTER)
                    # Empechements O/N
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GempEmpechement')))
                    wd.find_element(By.ID, 'inputB33ginf2Ya33GempEmpechement').send_keys(Keys.ENTER)
                    # Montant
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GmtMontant')))
                    wd.find_element(By.ID, 'inputB33ginf2Ya33GmtMontant').send_keys(Keys.ENTER)
                    # Date d'effet
                    # Jour
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour')))
                    wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetJour').send_keys(Keys.ENTER)
                    # Mois
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois')))
                    wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetMois').send_keys(Keys.ENTER)
                    # Année
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee')))
                    wd.find_element(By.ID, 'inputB33ginf2Ya33GdtefDateEffetAnnee').send_keys(Keys.ENTER)
                    # Ref jugt validite
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite')))
                    # wd.find_element(By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite').send_keys(data[j][4])
                    wd.find_element(By.ID, 'inputB33ginf2Ya33GjuvlJugementValidite').send_keys(Keys.ENTER)
                    # date d'execution jugt
                    # Jour
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementJour')))
                    wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementJour').send_keys(Keys.ENTER)
                    # Mois
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementMois')))
                    wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementMois').send_keys(Keys.ENTER)
                    # Année
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementAnnee')))
                    wd.find_element(By.ID, 'inputB33ginf2Ya33GdjuvDateExecutionJugementAnnee').send_keys(Keys.ENTER)
                    self.Deb("pas 83")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Date de renouvellement
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementJour')))
                    wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementJour').send_keys(Keys.ENTER)
                    time.sleep(delay)
                    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementMois')))
                    wd.find_element(By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementMois').send_keys(Keys.ENTER)
                    time.sleep(delay)
                    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf2Ya33GdtreDateRenouvellementAnnee')))
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
                    self.Deb("pas 84")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

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
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec')))
                    wd.find_element(By.ID, 'inputB33gsuprYa33G007ReponseSuitePrec').send_keys('S')
                    self.Deb("pas 85")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir REF main levée
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 30).until(EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GrfmlRefMainlevee')))
                    time.sleep(delay)
                    numero_ope = data[j][14]
                    wd.find_element(By.ID, 'inputB33ginf3Ya33GrfmlRefMainlevee').send_keys(f"{numero_ope} du {djc_capture}")
                    time.sleep(delay)
                    wd.find_element(By.ID, 'inputB33ginf3Ya33GrfmlRefMainlevee').send_keys(Keys.ENTER)
                    self.Deb("pas 86")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir date de main levée
                # Capture et réutilisation de la date journée comptable
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GdtmlDateMainleveeJour')))
                    wd.find_element(By.ID, 'inputB33ginf3Ya33GdtmlDateMainleveeJour').send_keys(djc[0])
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GdtmlDateMainleveeMois')))
                    wd.find_element(By.ID, 'inputB33ginf3Ya33GdtmlDateMainleveeMois').send_keys(djc[1])
                    time.sleep(delay)
                    WebDriverWait(wd, 40).until(EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GdtmlDateMainleveeAnnee')))
                    wd.find_element(By.ID, 'inputB33ginf3Ya33GdtmlDateMainleveeAnnee').send_keys(djc[2])
                    wd.find_element(By.ID, 'inputB33ginf3Ya33GdtmlDateMainleveeAnnee').send_keys(Keys.ENTER)
                    self.Deb("pas 87")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir type de main levée
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GtpmlTypeMainlevee')))
                    wd.find_element(By.ID, 'inputB33ginf3Ya33GtpmlTypeMainlevee').send_keys('TOTALE')
                    time.sleep(delay)
                    wd.find_element(By.ID, 'inputB33ginf3Ya33GtpmlTypeMainlevee').send_keys(Keys.ENTER)
                    self.Deb("pas 87")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # REF JUGT NULLITE
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GrfnlRefJugtNullite')))
                    wd.find_element(By.ID, 'inputB33ginf3Ya33GrfnlRefJugtNullite').send_keys(Keys.ENTER)
                    self.Deb("pas 88")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # DATE JUGT NULLITE
                # Jour
                try:
                    time.sleep(delay)
                    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GdtnlDateJugtNulliteJour')))
                    wd.find_element(By.ID, 'inputB33ginf3Ya33GdtnlDateJugtNulliteJour').send_keys(Keys.ENTER)
                    time.sleep(delay)
                    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GdtnlDateJugtNulliteMois')))
                    wd.find_element(By.ID, 'inputB33ginf3Ya33GdtnlDateJugtNulliteMois').send_keys(Keys.ENTER)
                    time.sleep(delay)
                    WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33ginf3Ya33GdtnlDateJugtNulliteAnnee')))
                    wd.find_element(By.ID, 'inputB33ginf3Ya33GdtnlDateJugtNulliteAnnee').send_keys(Keys.ENTER)
                    self.Deb("pas 89")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Saisir O pour valider (Voulez-vous valider ?)
                try:
                    time.sleep(delay)
                    validation_script = \
                        f'''document.getElementById('inputB33gvlcrYa33GvalcValidationCreation').setAttribute('value','O');'''
                    wd.execute_script(validation_script)
                    # WebDriverWait(wd, 20).until(EC.presence_of_element_located((By.ID, 'inputB33gvlcrYa33GvalcValidationCreation')))
                    time.sleep(delay)
                    wd.find_element(By.ID, 'inputB33gvlcrYa33GvalcValidationCreation').send_keys(Keys.ENTER)
                    self.Deb("pas 90")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

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
                    self.Deb("pas 91")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # fin de transaction :
                try:
                    time.sleep(delay)
                    fin_transaction_script = \
                        f'''document.getElementById('inputB33gcrmdYa33G012Reponse').setAttribute('value','N');'''
                    wd.execute_script(fin_transaction_script)
                    self.Deb("pas 92")
                except (TimeoutException, StaleElementReferenceException, ElementNotInteractableException):
                    # messagebox.showinfo("Service Interrompu !", message_service_interrompu)
                    # print("data", data[j])
                    # sav.saved_file(filename=saved_file, j=j, data=data, rep=sortie_repertoire,
                    #                columns=columns_sortie)
                    WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                    wd.find_element(By.ID, 'barre_outils:touche_f2').click()
                    wd.quit()

                # Retour au menu
                WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
                wd.find_element(By.ID, 'barre_outils:touche_f2').click()

                # convertir les colonnes numériques de la liste data en entier

                ## Marquage tâche faîte dans le fichier
                # match os.path.isfile(filepath1):
                #     case True:
                #         data[j][13] = f"{numero_ope} du {djc_capture}"
                #         # data[j][14] = numero_ope
                #         data[j][16] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                #         data[j][15] = success
                #         print("inscription des données dans la liste ligne 2991", data)
                #     case False:
                #         data[j][13] = f"{numero_ope} du {djc_capture}"
                #         data[j][3] = str(date_d_effet.strftime('%Y-%m-%d'))
                #         # data[j][14] = numero_ope
                #         data[j][16] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                #         data[j][15] = success
                data[j][16] = f"{numero_ope} du {djc_capture}"
                # data[j][3] = str(date_d_effet.strftime('%Y-%m-%d'))
                # data[j][14] = numero_ope
                data[j][17] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                data[j][18] = 'V'
                print("inscription des données ligne 1830", data)
                print("le N° de ligne est à la ligne 1831:", j)

                ## Incrementation ProgressBar
                # pb['value'] += 90 / nb_ligne
                # progressbar_label.destroy()
                # tab6.update()
                # progressbar_label = Label(tab6,
                #                           text=f"Le travail est en cours: {pb['value']:.2f}%  ~  "
                #                                f"L'opération de création de SATD sera terminé à  {heure_de_fin}")
                # progressbar_label.place(x=250, y=label_y)
                # pb.update()
                # tab6.update()
                print("les données à la nouvelle ligne : ", data[j])
                data_ods = data
                sheet_name = "Feuille1"
                data_ods.insert(0, columns_sortie)
                print("les nouvelles data 2725: \n", data)
                if data[0] != columns_sortie:
                    data.insert(0, columns_sortie)
                save_data(sortie_repertoire + fichier_de_Sortie, data)
                if data[0] == columns_sortie:
                    del data[0]
                end = time.time()
                print("Temps d'une boucle :", format(end - start))
                if j < nb_ligne - 1:
                    j += 1
                else:
                    break

            data_df = pd.DataFrame(data_ods)

            print("le dataframe : ", data_df)

            # try:
            #     time.sleep(delay)
            #     time.sleep(delay)
            #     time.sleep(delay)
            #     tabControl.add(tab4, text='Liste SATD en sortie de traitement')
            #     table1 = Table(tab4, dataframe=data_df, read_only=True, index=FALSE)
            #     table1.place(y=120)
            #     table1.autoResizeColumns()
            #     table1.show()
            #
            # except FileNotFoundError as e:
            #     print(e)
            #     messagebox.showerror('Erreur de tableau', 'Il n\'y a pas de tableau à afficher')
            # progressbar_label.destroy()
            # tab2.update()
            # progressbar_label = Label(tab2,
            #                           text=f"Le travail est maintenant fini! A bientôt")
            # progressbar_label.place(x=250, y=label_y)
            messagebox.showinfo("Données Manquante", "Le travail est maintenant fini! A bientôt")
            wd.quit()

    def Deb(self, msg):
        print(f"Debug  {msg if msg is not None else ''}-:{sys._getframe().f_back.f_lineno}")
