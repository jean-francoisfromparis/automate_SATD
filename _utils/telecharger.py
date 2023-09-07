import time
from tkinter import messagebox

from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class Telecharger_fichier:
    def telecharger(self,wd,liste_jour_a_telecharger):
        # Récupération des lignes du tableau du calendrier
        element_calendrier = wd.find_elements(By.CLASS_NAME,"day")
        calendrier = []
        for jour_calendrier in element_calendrier:
            calendrier.append(jour_calendrier.text)

        print(calendrier)
        indices_jour = []
        boutons_jour = [wd.find_element(By.CSS_SELECTOR,'#calendrier > div > table > tbody > tr:nth-child(2) > td.day.selected.today')]
        for n in range(len(liste_jour_a_telecharger)):
            # print()
            indices_jour.append(calendrier.index(str(liste_jour_a_telecharger[n].day)))
            # print(indice_jour)
            boutons_jour.append(element_calendrier[indices_jour[n]])
            # print(boutons_jour[n].text)

            try:
                filtre = wd.find_element(By.ID, 'filtre_id')
                filtre.clear()
            except:#any exception
                pass

            try:
                day_button = boutons_jour[n]
                time.sleep(10)
                filtre = wd.find_element(By.ID, 'filtre_id')
                filtre.send_keys("Credit_TVA")
                day_button.click()
                print("pas 1 - telecharger - ligne 13")
            except:  # any exception
                pass

            try:
                filtre_button = wd.find_element(By.CSS_SELECTOR, '#monmenu > ul > li:nth-child(5) > input[type=button]:nth-child(2)')
                time.sleep(10)
                filtre_button.click()
                print("pas 2 telecharger - ligne 20")
            except:
                pass

            try:
                dossiers = wd.find_elements(By.XPATH, '//*[starts-with(@id,"ico")]')
                # if len(dossiers) == 0:
                #     messagebox.showinfo('Pas de dossier Crédit de TVA', 'Il n\'y a pas de dossier '
                #                                                         'Crédit TVA aujourd\'hui à afficher')
                for i in range(len(dossiers)):
                    dossiers[i].click()
                print(len(dossiers))
                print("pas 3 telecharger - ligne 32")
            except:  # any exception
                pass

            try:
                WebDriverWait(wd, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, '#monmenu > ul > li:nth-child(6)')))
                tout_cocher = wd.find_element(By.CSS_SELECTOR, '#monmenu > ul > li:nth-child(6)')
                tout_cocher.click()
                print("pas 4 telecharger - ligne 39")
            except:  # any exception
                pass

            try:
                telecharger = wd.find_element(By.CSS_SELECTOR, 'li.titre:nth-child(7)')
                telecharger.click()
                time.sleep(10)
                print("pas 5 telecharger - ligne 46")
            except:  # any exception
                pass

            try:
                wd.switch_to.alert.accept()
                time.sleep(10)
            except:#any exception
                pass


            # try:
            #     fermer = wd.find_element(By.CSS_SELECTOR, 'a[href="../delogue.php"]')
            #     fermer.click()
            #     print("pas 6 telecharger - ligne 100")
            # except:#any exception
            #     pass
