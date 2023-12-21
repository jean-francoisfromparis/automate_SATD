import os
from datetime import datetime
import pandas as pd
from pyexcel_ods import save_data
class Data_verification:
    def data_verification(self,entree_df):
        # création du repértoire de sortie et du fichier de sortie avant traitement
        source_rep = os.getcwd()
        fichier_de_sortie = 'donnees_sortie_phaseII' + datetime.now().strftime('_%Y-%m-%d') + '.ods'
        repertoire_de_sortie = source_rep + '\donnees_sortie_phaseII\donnees_sortie_phaseII' + datetime.now().strftime(
            '_%Y-%m-%d')
        chemin_fichier_de_sortie = repertoire_de_sortie + '\\' + fichier_de_sortie
        print("-------------------------------------- Récupération des données --------------------------------------")
        # Vérification de l'existence d'un fichier de sortie à la date du jour
        # Création des colonnes pour la phaseII
        columns = list(entree_df)
        columns_sortie = columns + ["N°Opération Phase 2", "Date Opération phase 2", "Dossiers traités"]
        if not os.path.exists(repertoire_de_sortie):
            os.makedirs(repertoire_de_sortie)
        elif not os.path.exists(chemin_fichier_de_sortie):
            sortie_df = pd.DataFrame(columns=columns_sortie)
        else:
            print(chemin_fichier_de_sortie)
            sortie_df = pd.read_excel(chemin_fichier_de_sortie)
        print("Le fichier d'entrée contient " + str(entree_df.shape[0]) + " lignes.")
        sortie_df = entree_df
        sortie = sortie_df.values.tolist()
        sortie.insert(0, columns_sortie)
        save_data(chemin_fichier_de_sortie, sortie)
        print("Le fichier de sortie contient " + str(
            sortie_df.shape[0]) + "Lignes de la précédente opération.")
        print(sortie_df)

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
            "Numéro de facture Chorus"].astype(str, errors='ignore')
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
        # entree_df.insert(13, "N°Opération Phase 1", "0", allow_duplicates=False)
        entree_df["N°Opération Phase 1"] = entree_df["N°Opération Phase 1"].astype(str)
        # entree_df["Date Opération phase 1"] = entree_df["Date Opération phase 1"].astype(str)
