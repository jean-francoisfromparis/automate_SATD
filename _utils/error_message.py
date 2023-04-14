from datetime import time
from telnetlib import EC
from tkinter import messagebox

from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait


class ErrorMessage:
    def error_message(self,wd):
        delay = 3
        # WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.CLASS_NAME, 'ui-messages-error')))
        # messages = wd.find_element(By.CLASS_NAME, 'ui-messages-error').text
        WebDriverWait(wd, 60).until(EC.presence_of_element_located((By.ID, 'messages')))
        messages = wd.find_element(By.ID, 'messages').text
        print(messages)
        time.sleep(delay)
        time.sleep(delay)
        messagebox.showinfo("Service Interrompu !", messages)
        WebDriverWait(wd, 100).until(EC.presence_of_element_located((By.ID, 'barre_outils:touche_f2')))
        wd.find_element(By.ID, 'barre_outils:touche_f2').click()
        wd.close()
