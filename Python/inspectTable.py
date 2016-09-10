#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""inspectTable.py: Inspection automatique du contenu de documents Word : vérification des tableaux"""

__author__ = "Nicolas Dugué, IUT GEA, Université du Maine, LIUM"

from docx import Document

#La conversion EMU vers cm est un peu imparfaite donc on utilise une égalité à 0.01 cm près
def isclose (value, truth):
	return abs(value - truth) <= 0.01

#On boucle sur toutes les cellules et on regarde si ce qu'on a correspond à truth avec isclose 
def check(cells, truth):
	isOkay=True
	for idx,c in enumerate(cells):
		value=c.width/cm
		#print str(value)+"cm"
		okay=isclose(value, truth[idx])
		#print str(okay)
		isOkay&=okay
	return isOkay

#Ouverture du document
document = Document('Tableaux.docx')
#On récupère les cellules de la première ligne du premier tableau de ce document
cells=document.tables[0].rows[0].cells
#1cm=360000 EMU, unité **** utilisée dans docx
cm=360000.0
#Les valeurs que l'on doit trouver pour les largeurs de cellule
truth=[8,5,4]
isOkay=check(cells, truth)
print isOkay

#avec des valeurs de truth qui sont mauvaises, cela donne false
truth=[8,6,4]
isOkay=check(cells, truth)
print isOkay
