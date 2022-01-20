#!/usr/bin/python3


import xlrd
import string
import time


"""Pour importer les données excel"""
	

"""Extraction des données dans un fichier excel"""
book = xlrd.open_workbook("taf.xls")
sh = book.sheet_by_index(0)


# --------------------------------------- HTML / CSS -------------------------------------------

def write_html_header(site, titre):
    site.write("""
   <!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>Météo Aéroport</title>
    <link href="style.css" rel="stylesheet" type="text/css">
</head>

    <body>

    <header>
    <h1><b>Météorologie</b></h1>
</header>

<section id="entete">
    <h3 id="aeroport">LFRS - Aéroport Nantes Atlantique </h3>
</section>

    """)


def write_html_nav_begin(site):
    site.write("<nav>\n")

def write_html_nav_end(site):
    site.write("</nav>\n")

    
def write_html_nav(site, date_complete, date_complete1, date_complete2):
    site.write("<ul>")
    site.write("""<li><a href="./site_page_suivante.html#entete" id="ici">""")
    site.write("{0}".format(date_complete2))
    site.write("</a></li>")
    site.write("""<li><a href="./site.html#s2">""")
    site.write("{0}".format(date_complete))
    site.write("</a></li>")
    site.write("""<li><a class="active" href="./site_page_precedente.html#entete"> """)
    site.write("{0}".format(date_complete1)) 	
    site.write("</a></li>")	
    site.write("</ul>")
 

def write_html_body_begin(site):
    site.write("<body>\n")

def write_html_body_end(site):
    site.write("</body>\n")


def write_html_body(site):
    site.write("""

<style>
body {
    background-color: #8c8c8c;
    background-repeat: no-repeat;
    background-size: cover;
}

header {
    background-color: #4d4d4d;
    background-attachment: fixed;
    height: 660px;
    border-style: solid;
    border-color: #000000;
}

h1 {
    margin: 15%;
    font-size: 170px;
}

#entete {
    background-color: #0e2bb7;
    height: 70px;
    margin: 0px;
    padding-left: 10px;
    border-radius: 10px;
    text-align: center;
    padding-top: 0. 5%;
}

#aeroport {
    color: #000000;
    font-size: 30px;
    padding-top: 0.8%;
    margin: 0px;
}


nav {
    height: 60px;
    width: 1000px;
    position: relative;
    margin-top: 50px;
    margin-left: 19%;
    margin-bottom: 40px;
}

ul {
    list-style-type: none;
    margin: 0;
    padding-right: 56px;
    padding-left: 0px;
    overflow: hidden;
    background-color: #485dd3;
    height: 100%;
    border-radius: 50px;
}

li {
    float: right;
}

li a {
    display: block;
    color: white;
    text-align: center;
    height: 100%;
    width: 280px;
    padding: 15px 10px 20px;
    text-decoration: none;
    font-size: large;
}

li a:hover {
    background-color: #3b4ba6;
    color: #000000;
    border: solid;
}

#ici {
	background-color: blue;
	color : white
}



#s2 {
    height: 490px;
    width: 46%;
    padding: 1%;
    padding-top: 0%;
    margin-top: 10px;
    margin-left: 15px;
    background: #6b6bb9;
    border-radius: 10px;
    text-align: justify;
    float: left;
    margin-bottom: 15px
}

#s3 {
    height: 490px;
    width: 46%;
    padding: 1%;
    padding-top: 0%;
    margin-top: 10px;
    background: #6b6bb9;
    border-radius: 10px;
    text-align: justify;
    float: left;
    margin-left: 20px;
    margin-bottom: 15px;
}

h4 {
    text-align: center;
}



     /* <---------------- PAGE JOUR J+1  ----------------> */


#prevision_j1 {
    height: 490px;
    width: 100%;
    text-align: center;
    padding: 0px 20% 0px 20%;
}

#div_prevision {
    height: auto;
    width: 61%;
    text-align: center;
    padding: 1%;
    padding-top: 0%;
    margin-top: 10px;
    background: #6b6bb9;
    border-radius: 10px;
    margin-left: 20px;


}

#titre_prevision {
    text-align: center;
    margin: 0px;
    padding-top: 20px;
}
</style>
""")
	
    


"""SECTION"""
def write_html_s3_begin(site):
    site.write("<section id=""prevision_j1"">\n")

def write_html_s3_end(site):
    site.write("</section>\n")

    
def write_html_s3(site):
    site.write("")
    
    
""" DIV """
def write_html_s3_div1_begin(site):
    site.write("<div id=""div_prevision"">\n")

def write_html_s3_div1_end(site):
    site.write("</div>\n")

    
def write_html_s3_div1(site):
    site.write("<h4 id=""id=titre_prevision"">")
    site.write("Prévision sur la journée")
    site.write("</h4>")

def write_html_end(site):
    site.write("</html>")
    
    
    
    
# --------------------------- Python - Définition de variable --------------------------------

			
#"""Prévision"""

def date() :
	fonction_date = format(sh.cell_value(rowx=1, colx=1))
	date2 = ""

	def checkInt(str):
    		try:
        		int(str)
        		return True
    		except ValueError:
        		return False

	for car in fonction_date:
    		if not checkInt(car):
        		date2+=car

	text = date2
	date, sep, tail = text.partition(':')
	
	y3 = 1
	ds = format(sh.cell_value(rowx=1, colx=y3))[:5]
	
	
	
	cell = format(sh.cell_value(rowx=1, colx=y3))
	text = cell.lstrip(string.ascii_letters)
	heure_tableau, sep, tail = text.partition('>')

	while ds == date :
		if ds == date : 
			y3 = y3+1 
			ds = format(sh.cell_value(rowx=1, colx=y3))[:5]
			
			cell = format(sh.cell_value(rowx=1, colx=h))
			text = cell.lstrip(string.ascii_letters)
			heure_tableau, sep, tail = text.partition('>')
			

	fonction_date2 = format(sh.cell_value(rowx=1, colx=y3))
	date22 = ""

	def checkInt(str):
    		try:
        		int(str)
        		return True
    		except ValueError:
        		return False

	for car in fonction_date2:
    		if not checkInt(car):
        		date22+=car

	text = date22
	date2, sep, tail = text.partition(':')
	
	site.write("<br> <b>Pour {0} ".format(date2))
	site.write(" à {0} :</b>" .format(heure_tableau))
	site.write("<br> La vitesse du vent est de {0} <br>".format(sh.cell_value(rowx=7, colx=y3)))
	site.write("L'orientation du vent est de {0} <br>".format(sh.cell_value(rowx=6, colx=y3)))
	site.write("La visibilité est de {0} <br>".format(sh.cell_value(rowx=4, colx=y3)))
	site.write("La météo est {0} ".format(sh.cell_value(rowx=3, colx=y3)))
	site.write("à {0} <br>".format(sh.cell_value(rowx=5, colx=y3)))
	
	a1= "Les rafales sont de {0} <br>".format(sh.cell_value(rowx=8, colx=y3))
	a2= "Il n'y a pas de Rafales <br>"
	if format(sh.cell_value(rowx=8, colx=y3)) == True :
        	site.write(a1)
	else : 
        	site.write(a2) 
        	
	soleil1 = str(sh.cell_value(rowx=9, colx=y3)[:2]) 
	if soleil1 == True and soleil1 < str(10) :
        	site.write("Le levé du soleil est à {0} <br>".format(sh.cell_value(rowx=9, colx=y3)))
	elif soleil1 == True and soleil1 > str(10) :
        	site.write("Le couché du soleil est à {0} <br>".format(sh.cell_value(rowx=9, colx=y3)))
	else :
		site.write("<br>")
		

def date2() :
	
	fonction_date = format(sh.cell_value(rowx=1, colx=1))
	date2 = ""

	def checkInt(str):
    		try:
        		int(str)
        		return True
    		except ValueError:
        		return False

	for car in fonction_date:
    		if not checkInt(car):
        		date2+=car

	text = date2
	date, sep, tail = text.partition(':')
	
	y3 = 1
	ds2 = format(sh.cell_value(rowx=1, colx=y3))[:5]

	while ds2 == date :
		if ds2 == date : 
			y3 = y3+1 
			ds2 = format(sh.cell_value(rowx=1, colx=y3))[:5]
			
			h2 = y3
			cell = format(sh.cell_value(rowx=1, colx=h2))
			text = cell.lstrip(string.ascii_letters)
			heure_tableau, sep, tail = text.partition('>')
	
	ds4 = ds2[:4]	
	y4 = y3
	ds3 = format(sh.cell_value(rowx=1, colx=y3))[:4]
	
	
	while ds4 == ds3 :
		
		
		if ds4 == ds3 :
			y4 = y4+1 
			ds4 = format(sh.cell_value(rowx=1, colx=y4))[:4]
			ds6 = ds4[:4]
			
			h4 = y4
			cell = format(sh.cell_value(rowx=1, colx=h4))
			text = cell.lstrip(string.ascii_letters)
			heure_tableau, sep, tail = text.partition('>')
		
			
			fonction_date40 = format(sh.cell_value(rowx=1, colx=h4))
			date40 = ""

			def checkInt(str):
    				try:
        				int(str)
        				return True
    				except ValueError:
        				return False

			for car in fonction_date40:
    				if not checkInt(car):
        				date40+=car

			text = date40
			date4, sep, tail = text.partition(':')	
			
			site.write("<br> <b>Pour {0} ".format(date4))
			site.write(" à {0} :</b>" .format(heure_tableau[:5]))
			site.write("<br> La vitesse du vent est de {0} <br>".format(sh.cell_value(rowx=7, colx=y4)))
			site.write("L'orientation du vent est de {0} <br>".format(sh.cell_value(rowx=6, colx=y4)))
			site.write("La visibilité est de {0} <br>".format(sh.cell_value(rowx=4, colx=y4)))
			site.write("La météo est {0} ".format(sh.cell_value(rowx=3, colx=y4)))
			site.write("à {0} <br>".format(sh.cell_value(rowx=5, colx=y4)))
			
			a1= "Les rafales sont de {0} <br>".format(sh.cell_value(rowx=8, colx=y4))
			a2= "Il n'y a pas de Rafales <br>"
			
			if format(sh.cell_value(rowx=8, colx=y4)) == True :
        			site.write(a1)
			else : 
        			site.write(a2) 
        	
			soleil1 = str(sh.cell_value(rowx=9, colx=y4)[:2]) 
			
			if soleil1 == True and soleil1 < str(10) :
        			site.write("Le levé du soleil est à {0} <br>".format(sh.cell_value(rowx=9, colx=y4)))
			elif soleil1 == True and soleil1 > str(10) :
        			site.write("Le couché du soleil est à {0} <br>".format(sh.cell_value(rowx=9, colx=y4)))

			else :
				site.write("<br>")
	
		

# ----------------------------- Python - Appelle des variables -------------------------------

import datetime
from datetime import datetime, timedelta

site = open("site_page_suivante.html", "w")
write_html_header(site,"Mon titre")
write_html_nav_begin(site)

date_complete = str("Actuellement")
date_complete1 = str("Heures précédentes")
date_complete2 = str("Toute la journée")

write_html_nav(site, date_complete, date_complete1, date_complete2)
write_html_nav_end(site)
write_html_body_begin(site)
write_html_body(site)
write_html_s3_begin(site)
write_html_s3_div1_begin(site)
write_html_s3_div1(site)

date()
date2()

write_html_s3_div1_end(site)
write_html_s3_end(site)
write_html_body_end(site)
write_html_end(site)

site.close()
