# -*- coding: utf-8 -*-
"""
Created on Mon Nov 18 12:53:03 2024

@author: SAIM Bahaeddine
"""
import pandas as pd
import pywhatkit as pwk
import time

# Charger le fichier Excel
df = pd.read_excel("Classeur2.xlsx")

# Message à envoyer
message = """
Chères et chers étudiant(e)s,

Nous sommes heureux de vous confirmer la date de notre rencontre pour l’entretien et la réunion de bienvenue du Club Sportif EHEI.

Détails de l’événement :

Date : Mercredi 13 novembre 2024
Heure : 13h10
Lieu : Amphithéâtre 1
Cette réunion sera l'occasion de mieux vous connaître et de partager avec vous les activités et projets de notre club. Merci de bien vouloir arriver quelques minutes à l’avance pour garantir le bon déroulement de la session.

Pour toute question ou précision, n’hésitez pas à nous contacter par téléphone au +212 6 ---------.

Nous nous réjouissons de vous accueillir et de débuter cette aventure sportive ensemble !

Bien cordialement,
"""

# Envoi du message à chaque contact
for index, row in df.iterrows():
    numero = str(row['Numero'])  # Assurez-vous que la colonne est bien nommée "Numero"
    
    # Vérification et formatage du numéro avec le préfixe +212
    if numero.startswith("0"):
        numero = "+212" + numero[1:]
    elif numero.startswith("6") or numero.startswith("7"):
        numero = "+212" + numero
    elif not numero.startswith("+212"):
        numero = "" + numero

    try:
        # Envoyer le message
        pwk.sendwhatmsg_instantly(numero, message, wait_time=10, tab_close=True)
        print(f"Message envoyé à {numero}")
        time.sleep(2)  # Pause pour éviter des blocages
    except Exception as e:
        print(f"Erreur lors de l'envoi à {numero}: {e}")


