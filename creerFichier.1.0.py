##################################################################
##	Creation de nouveau fichier comptes Arbre à Poule		  	##
##################################################################
##	Ce programme permet de creer un nouveau fichier de comptes	##
##	pour le mois suivant, à partir du fichier vide original	  	##
##	et de la base de prix du mois précédemment utilisé			##
##																##
##################################################################
##	Développeur: 	Henri DEWILDE								##
##					henri.dewilde@gmail.com 					##
##################################################################
##  Version 1  :    decembre 2016								##
##################################################################

import socket
import os
import sys

sys.path.append('utils/')
import pyoocalc
import uno

	################################################################
	### 	définition des variables
	################################################################

liste_nom = ("01 janvier", 
			"02 fevrier",
			"03 mars",
			"04 avril",
			"05 mai",
			"06 juin",
			"07 juillet",
			"08 aout",
			"09 septembre",
			"10 octobre",
			"11 novembre",
			"12 decembre")


	################################################################
	### 	définition des fonctions
	################################################################

# fonction qui extrait les valeurs numeriques d'une liste et les renvoie
def liste_num (liste0):
	liste1 = []

	for val in liste0 :
		if val.isdigit() :
			liste1.append(val)
	#print(liste1)
	return liste1

# fonction qui extrait les noms de fichiers correspondants aux mois de l'année
def extr_mois (liste0):
	liste1 = []

	for val in liste0 :
		if val[:2].isdigit() :
			liste1.append(val)
	#print (liste1)
	return liste1

# fonction qui extrait les noms de fichiers correspondants aux fichiers vides
def extr_fich_vide(liste0):
	liste1 = []

	for val in liste0 :
		if val[:7] == "comptes" :
			liste1.append(val)
	return liste1

# fonction qui donne le nom du nouveau fichier à enregistrer
def nouv_nom (nom):
	nouveau_nom = ""

	if nom[:2] == "12" :
		return liste_nom [0]
		
	else :
		i = 0

		while nom[:2] != liste_nom [i][:2] :
			i += 1

		return liste_nom [i+1]


	################################################################
	### 	Corp du programme
	################################################################

	################################################################
	# recherche du répertoire de travail
	################################################################
print ("## recherche du répertoire de travail")

# On liste les dossiers
dossiers = os.listdir (os.getcwd()+"/../")

dossiers_num = liste_num(dossiers)

print ( "dernière année renseignée : " + max(dossiers_num) )

annee = max(dossiers_num)

#continuer = input ( " continuer? " )	# !!!!


	################################################################
	# Ouverture du dernier fichier utilisé 
	################################################################
print ("## recherche du dernier fichier utilisé")

#recherche du dernier fichier utilisé
fichiers = os.listdir (os.getcwd() + "/../" + str(annee))
#print ( fichiers )

fichiers_ok = extr_mois (fichiers)

dern_fich = max (fichiers_ok)	#dernier fichier utilisé

print ("dernier fichier : " + dern_fich)

print ("## ouverture du dernier fichier utilisé")

# open document
doc = pyoocalc.Document()
file_name = os.getcwd() + "/../" + str(annee) + "/" + dern_fich
doc.open_document(file_name)

# access the active sheet
sheet1 = doc.sheets.sheet(0)

produits = {}
prix = {}

i = 3

print ("## enregistrement de la base de prix du dernier fichier utilisé")

# copie de la liste des produits et prix dans une table
while i <= 200 : 

	if sheet1.cell_value_by_index(0, i) != None :

		produits [i] = sheet1.cell_value_by_index(0, i)
		prix [i] = sheet1.cell_value_by_index(1, i) 

		print ("produit : " + str(produits [i]))
		print ("prix : " + str(prix[i]) )
	else :
		produits [i] = ""
		prix [i] = ""
		
	i += 1

print ("## fermeture du dernier fichier utilisé")
doc.close_document()

	################################################################	
	# Ouverture du fichier vide 
	################################################################
print ("## création du nouveau fichier")

#recherche de la dernière version du fichier vide
fichiers = os.listdir (os.getcwd()+"/../fichier vide original")
#print ( fichiers )

fichiers_ok = extr_fich_vide (fichiers)

nom_vide = max (fichiers_ok)	#dernière version du fichier vide
print ("utilisation du fichier vide original nommé : " + nom_vide)

# open document
doc = pyoocalc.Document()
file_name = os.getcwd() + "/../fichier vide original/" + nom_vide
doc.open_document(file_name)

sheet0 = doc.sheets.sheet(0)

print ("## alimentation base de prix du nouveau fichier")
# alimentation de la base de prix du fichier vide
i = 4
while i <= 200 :

	if produits [i] != "" :
		sheet0.set_cell_value_by_index(produits [i], 0, i)
		sheet0.set_cell_value_by_index(prix [i], 1, i)
		
	i += 1


	################################################################
	# sauvegarde et fermeture du document
	################################################################
print ("## sauvegarde du nouveau fichier")

# recherche du nom du fichier à créer
nom_nouveau = nouv_nom(dern_fich)

print ("nom nouveau fichier : " + nom_nouveau)

# Si on crée le mois de janvier
if nom_nouveau == liste_nom[0] :
	annee = int(annee) + 1
	os.makedirs("../" + str(annee))		# création du dossier de l'année suivante

# Sauvegarde du nouveau fichier
file_name = os.getcwd() + "/../" + str(annee) + "/" + nom_nouveau + ".ods"
doc.save_document(file_name)

print ("Le fichier " + nom_nouveau + ".ods a été créé avec succès dans le dossier " + str(annee))

doc.close_document()


	################################################################
	### Fin du programme
	################################################################
print ("## fin normale du programme")
#input ("")