from datetime import time, datetime
from tkinter import messagebox

from selenium.common import TimeoutException, StaleElementReferenceException, ElementNotInteractableException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from _utils.save_file import Saved_file


class Error_correction:
    def error_correction(self,delay,wd,chemin_fichier_de_sortie,j,data, repertoire_de_sortie,columns_sortie,message_service_interrompu):
        sav = Saved_file()
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

        # Validation de la non saisie des dates
        try:
            time.sleep(delay)
            WebDriverWait(wd, 40).until(
                EC.presence_of_element_located((By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON')))
            wd.find_element(By.ID, 'inputBrep9081Rep9082ReponseUtilisateurON').send_keys('O')
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