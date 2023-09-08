import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


class Telecharger_fichier:
    def telecharger(self,headless,liste_jour_a_telecharger,delay):
        login = "meddb-jean-francois.consultant"
        mot_de_passe = "Dagobert01"
        print("la liste de jour:",liste_jour_a_telecharger)
        for n in range(len(liste_jour_a_telecharger)):
            print("le jour en cours de téléchargement",liste_jour_a_telecharger[n])
            wd_options = Options()
            # wd_options.headless = headless
            if headless:
                wd_options.add_argument('-headless')

            wd_options.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'
            wd_options.set_preference('detach', True)
            wd_options.add_argument("--enable-javascript")
            wd = webdriver.Firefox(options=wd_options)
            # wd = webdriver.Firefox(executable_path=GeckoDriverManager().install(), options=wd_options)
            # TODO Passer au service object
            wd.get(
                'https://portailmetierpriv.appli.impots/cas/login?service=http://pdf-integ.appli.dgfip/login.php')  # adresse PDF EDIT
            # Elimination des onglet about-blank
            all_tab = wd.window_handles
            wd.switch_to.window(all_tab[0])
            time.sleep(delay)
            i = 0
            for i in range(len(all_tab)):
                wd.switch_to.window(all_tab[i])
                time.sleep(delay)
                time.sleep(delay)
                if not wd.title:
                    wd.close()
                elif wd.title == "Protection de la navigation par F-Secure":
                    print(wd.title)
                time.sleep(delay)
            new_tabs = wd.window_handles
            wd.switch_to.window(new_tabs[0])
            # Saisir utilisateur
            while wd.title == "Identification":
                print(wd.title)
                time.sleep(delay)
                wd.find_element(By.ID, 'identifiant').send_keys(login)
                wd.find_element(By.ID, 'identifiant').send_keys(Keys.TAB)
                # Saisie mot de pass
                time.sleep(delay)
                # wd.find_element(By.ID, 'secret_tmp').send_keys(mot_de_passe)
                wd.find_element(By.ID, 'secret_tmp').send_keys(mot_de_passe)

                time.sleep(delay)
                wd.find_element(By.ID, 'secret_tmp').send_keys(Keys.RETURN)
                time.sleep(delay)
            print(wd.title)

            # cliquer sur MDA
            try:
                if wd.title == "PDFEDIT - Consultation prog":
                    WebDriverWait(wd, 20).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 'choix')))
                    mda_button = wd.find_element(By.ID, 'MDA')
                    mda_button.click()
                    print("pas 1-ligne 918")
            except:
                pass

            try:
                dge_button = wd.find_element(By.CSS_SELECTOR,
                                             'body > div:nth-child(2) > p:nth-child(3) > a:nth-child(3214)')
                dge_button.click()
                wd.switch_to.default_content()
                frames = wd.find_elements(By.TAG_NAME, "frame")
                print("La liste de frame contient " + str(len(frames)))
                wd.switch_to.frame(frames[1])
                print("pas 2-ligne 944")
            except:
                pass
            # Récupération des lignes du tableau du calendrier
            element_calendrier = wd.find_elements(By.CLASS_NAME,"day")
            calendrier = []
            for jour_calendrier in element_calendrier:
                calendrier.append(jour_calendrier.text)

            print(calendrier)
            # boutons_jour = [wd.find_element(By.CSS_SELECTOR,'#calendrier > div > table > tbody > tr:nth-child(2) > td.day.selected.today')]
            print(str(liste_jour_a_telecharger[n].day))
            indices_jour = calendrier.index(str(liste_jour_a_telecharger[n].day))
                # print(indice_jour)
            boutons_jour = element_calendrier[indices_jour]
            print("le texte du bouton:",boutons_jour.text)

            try:
                filtre = wd.find_element(By.ID, 'filtre_id')
                filtre.clear()
            except:#any exception
                pass

            try:
                day_button = boutons_jour
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
                WebDriverWait(wd, 40).until(EC.element_to_be_clickable((By.XPATH, '/html/body/form[1]/div[1]/ul/li[6]')))
                time.sleep(10)
                tout_cocher = wd.find_element(By.XPATH, '/html/body/form[1]/div[1]/ul/li[6]')
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

            try:
                fermer = wd.find_element(By.CSS_SELECTOR, 'a[href="../delogue.php"]')
                fermer.click()
                print("pas 6 telecharger - ligne 90")
            except:#any exception
                pass

            wd.close()
