#!/usr/bin/python3

# """ Imporrtation """
import xlrd
import string
import time
import datetime
from datetime import datetime, timedelta

# ----------------------------  HTML / CSS  ----------------------------------------------
	

#""" Extraction des données dans un fichier excel """
book = xlrd.open_workbook("taf.xls")
sh = book.sheet_by_index(0)

#""" Header """
def write_html_header(site, titre):
    site.write("""
    <!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <title>Météo Aéroport</title>
    <link href="style.css" rel="stylesheet" type="text/css">
</head>

<header>
<h1><b>Météorologie</b></h1>
</header>

<section id="entete">
    <h3 id="aeroport">LFRS - Aéroport Nantes Atlantique </h3>
</section>

    """)

# """ Nav """
def write_html_nav_begin(site):
    site.write("<nav>\n")

def write_html_nav_end(site):
    site.write("</nav>\n")

    
def write_html_nav(site, date_complete, date_complete1, date_complete2):
    site.write("<ul>")
    site.write("""<li><a href="./site_page_suivante.html#entete">""")
    site.write("{0}".format(date_complete2))
    site.write("</a></li>")
    site.write("""<li><a href="./site.html#s2" id="ici">""")
    site.write("{0}".format(date_complete))
    site.write("</a></li>")
    site.write("""<li><a class="active" href="./site_page_precedente.html#entete"> """)
    site.write("{0}".format(date_complete1)) 	
    site.write("</a></li>")	
    site.write("</ul>")
 

# """ Body """
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
    text-align: center;
    float: left;
    margin-left: 20px;
    margin-bottom: 15px;
}

h4 {
    text-align: center;
}



     /* <---------------- PAGE JOUR J+1  ----------------> */


#prevision_j1 {
    height: auto;
    width: 100%;
    text-align: center;
    padding: 0px 20% 0px 20%;
}

#div_prevision {
    height: 490px;
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
	

#"""SECTION"""
def write_html_s2_begin(site):
    site.write("<section id=""s2"">\n")

def write_html_s2_end(site):
    site.write("</section>\n")

    
def write_html_s2(site):
    site.write("")
    
    

def write_html_s3_begin(site):
    site.write("<section id=""s3"">\n")

def write_html_s3_end(site):
    site.write("</section>\n")

    
def write_html_s3(site):
    site.write("")
    
    
#""" DIV """

def write_html_s2_div1_begin(site):
    site.write("<div id="">\n")

def write_html_s2_div1_end(site):
    site.write("</div>\n")

    
def write_html_s2_div1(site):
    site.write("<h4>")
    site.write("Observation")
    site.write("</h4>")

def write_html_end(site):
    site.write("</html>")


def write_html_s3_div2_begin(site):
    site.write("<div id="">\n")

def write_html_s3_div2_end(site):
    site.write("</div>\n")

    
def write_html_s3_div2(site):
    site.write("<h4>")
    site.write("Prévision sur 2 heures")
    site.write("</h4>")

def write_html_end(site):
    site.write("</html>")





# --------------------------------------- Python -------------------------------------------



#---------------------------------  Renan  ------------------------------------

document = xlrd.open_workbook("metars.xls")

feuille_1 = document.sheet_by_index(0)
feuille_1 = document.sheet_by_name("Table 0")

rows = feuille_1.nrows

A1 = feuille_1.cell_value(rowx=0, colx=0)
A2 = feuille_1.cell_value(rowx=1, colx=0)
A3 = feuille_1.cell_value(rowx=2, colx=0)
A4 = feuille_1.cell_value(rowx=3, colx=0)
A5 = feuille_1.cell_value(rowx=4, colx=0)
A6 = feuille_1.cell_value(rowx=5, colx=0)
A7 = feuille_1.cell_value(rowx=6, colx=0)
A8 = feuille_1.cell_value(rowx=7, colx=0)
A9 = feuille_1.cell_value(rowx=8, colx=0)
A10 = feuille_1.cell_value(rowx=9, colx=0)
A11 = feuille_1.cell_value(rowx=10, colx=0)
A12 = feuille_1.cell_value(rowx=11, colx=0)
A13 = feuille_1.cell_value(rowx=12, colx=0)
A15 = feuille_1.cell_value(rowx=14, colx=0)
A16 = feuille_1.cell_value(rowx=15, colx=0)
A17 = feuille_1.cell_value(rowx=16, colx=0)
A18 = feuille_1.cell_value(rowx=17, colx=0)

B1 = feuille_1.cell_value(rowx=0, colx=1)
B2 = feuille_1.cell_value(rowx=1, colx=1)
B3 = feuille_1.cell_value(rowx=2, colx=1)
B4 = feuille_1.cell_value(rowx=3, colx=1)
B5 = feuille_1.cell_value(rowx=4, colx=1)
B6 = feuille_1.cell_value(rowx=5, colx=1)
B7 = feuille_1.cell_value(rowx=6, colx=1)
B8 = feuille_1.cell_value(rowx=7, colx=1)
B9 = feuille_1.cell_value(rowx=8, colx=1)
B10 = feuille_1.cell_value(rowx=9, colx=1)
B11 = feuille_1.cell_value(rowx=10, colx=1)
B12 = feuille_1.cell_value(rowx=11, colx=1)
B13 = feuille_1.cell_value(rowx=12, colx=1)
B15 = feuille_1.cell_value(rowx=14, colx=1)
B16 = feuille_1.cell_value(rowx=15, colx=1)
B17 = feuille_1.cell_value(rowx=16, colx=1)
B18 = feuille_1.cell_value(rowx=17, colx=1)

X = []
Y= []
for r in range(1, rows):
    X += [feuille_1.cell_value(rowx=r, colx=0)]
    Y += [feuille_1.cell_value(rowx=r, colx=1)]





def write_html_table(site, colonne1, colonne2):
    site.write("<table>\n")
    site.write("<tr>")
    site.write("\t<th> Catégorie</td>")
    site.write("\t<th> Données </td>")
    site.write("\t</tr>\n")
    for i in range(len(colonne1)):
         site.write("<tr>")
         site.write("\t<td>"+str(colonne1[i])+"</td>")
         site.write("\t<td>"+str(colonne2[i])+"</td>")
         site.write("\t</tr>\n")
    site.write("</table>\n")



site = open("site.html", "w")
colonne1 = [A1, A2, A3, A4, A5, A6, A8, A9, A10, A11, A6, A15, A16, A17, A18]
colonne2 = [B1, B2, B3, B4, B5, B6, B8, B9, B10, B11, B6, B15, B16, B17, B18]



#----------------------------------  Youli  --------------------------------------

#Phrase du jour
def date_heure():

	heure_actuel = str(time.strftime('%Hh%M', time.localtime()))
	
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
	data = str(time.strftime('%d/%m/%Y', time.localtime()))
	
	site.write("Le bulletin météorologique pour ce <b>{0}</b>".format(date))
	site.write(" <b>{0}</b>".format(data))
	site.write(" a été effectué à <b>{0} </b><br><br>".format(heure_actuel))


def prevision_s() :

#PREVISION A HEURE+1
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
	
	
	heure_actuel = str(time.strftime('%Hh%M', time.localtime()))
	ha = str(time.strftime('%H:%M', time.localtime()))[:2]
	
	y = 1
	y1 = 1
	cell = format(sh.cell_value(rowx=1, colx=y))
	text = cell.lstrip(string.ascii_letters)
	heure_tableau, sep, tail = text.partition('>')
	
	ht = heure_tableau[:2]

	
	while ht != ha : 
		if ht != ha : 
			y = y+1
			cell = format(sh.cell_value(rowx=1, colx=y))
			text = cell.lstrip(string.ascii_letters)
			heure_tableau, sep, tail = text.partition('>')
			ht = heure_tableau[:2]
		elif ht == ha : 
			break	
	y1 = y
	
	if ht == ha : 	
		y1 = y1+1
		cell = format(sh.cell_value(rowx=1, colx=y1))
		text = cell.lstrip(string.ascii_letters)
		heure_tableau, sep, tail = text.partition('>')
		ht = heure_tableau[:2]
			
	
	site.write("<br> <b> Pour {0} ".format(date))
	site.write(" à {0} :</b>" .format(heure_tableau))
	site.write("<br> La vitesse du vent est de {0} <br>".format(sh.cell_value(rowx=7, colx=y1)))
	site.write("L'orientation du vent est de {0} <br>".format(sh.cell_value(rowx=6, colx=y1)))
	site.write("La visibilité est de {0} <br>".format(sh.cell_value(rowx=4, colx=y1)))
	site.write("La météo est {0} ".format(sh.cell_value(rowx=3, colx=y1)))
	site.write("à {0} <br>".format(sh.cell_value(rowx=5, colx=y1)))
	
	a1= "Les rafales sont de {0} <br>".format(sh.cell_value(rowx=8, colx=y1))
	a2= "Il n'y a pas de Rafales <br>"
	if format(sh.cell_value(rowx=8, colx=y1)) == True :
        	site.write(a1)
	else : 
        	site.write(a2) 
        	
	soleil1 = str(sh.cell_value(rowx=9, colx=y1)[:2]) 
	if soleil1 == True and soleil1 < str(10) :
        	site.write("Le levé du soleil est à {0} <br>".format(sh.cell_value(rowx=9, colx=y1)))
	elif soleil1 == True and soleil1 > str(10) :
        	site.write("Le couché du soleil est à {0} <br>".format(sh.cell_value(rowx=9, colx=y1)))
	else :
		site.write("<br>")
	
	
#BOUCLE PREVISION A HEURE+2
	fonction_date10 = format(sh.cell_value(rowx=1, colx=y1))
	date10 = ""
	
	def checkInt(str):
    		try:
        		int(str)
        		return True
    		except ValueError:
        		return False

	for car in fonction_date10:
    		if not checkInt(car):
        		date10+=car

	text = date10
	ds, sep, tail = text.partition(':')
	
	while ds == date :
		if ds == date : 
			y1 = y1+1
			ds = format(sh.cell_value(rowx=1, colx=y1))
			cell = format(sh.cell_value(rowx=1, colx=y1))
			text = cell.lstrip(string.ascii_letters)
			heure_tableau, sep, tail = text.partition('>')
			ht = heure_tableau[:2]
			
			fonction_date20 = format(sh.cell_value(rowx=1, colx=y1))
			date20 = ""
	
			def checkInt(str):
    				try:
        				int(str)
        				return True
    				except ValueError:
        				return False

			for car in fonction_date20:
    				if not checkInt(car):
        				date20+=car

			text = date20
			dt, sep, tail = text.partition(':')
			
			
			site.write("<br> <b> Pour {0} ".format(dt))
			site.write(" à {0} :</b>" .format(heure_tableau))
			site.write("<br> La vitesse du vent est de {0} <br>".format(sh.cell_value(rowx=7, colx=y1)))
			site.write("L'orientation du vent est de {0} <br>".format(sh.cell_value(rowx=6, colx=y1)))
			site.write("La visibilité est de {0} <br>".format(sh.cell_value(rowx=4, colx=y1)))
			site.write("La météo est {0} ".format(sh.cell_value(rowx=3, colx=y1)))
			site.write("à {0} <br>".format(sh.cell_value(rowx=5, colx=y1)))
	
			a1= "Les rafales sont de {0} <br>".format(sh.cell_value(rowx=8, colx=y1))
			a2= "Il n'y a pas de Rafales <br>"
			if format(sh.cell_value(rowx=8, colx=y1)) == True :
        			site.write(a1)
			else : 
        			site.write(a2) 
        	
			soleil1 = str(sh.cell_value(rowx=9, colx=y1)[:2]) 
			if soleil1 == True and soleil1 < str(10) :
        			site.write("Le levé du soleil est à {0} <br>".format(sh.cell_value(rowx=9, colx=y1)))
			elif soleil1 == True and soleil1 > str(10) :
        			site.write("Le couché du soleil est à {0} <br>".format(sh.cell_value(rowx=9, colx=y1)))
			else :
				site.write("<br>")
		else :
			break	




site = open("site.html", "w")
write_html_header(site,"Mon titre")
write_html_nav_begin(site)

date_complete = str("Actuellement")
date_complete1 = str("Heures précédentes")
date_complete2 = str("Toute la journée")

write_html_nav(site, date_complete, date_complete1, date_complete2)
write_html_nav_end(site)
write_html_body_begin(site)
write_html_body(site)

write_html_s2_begin(site)
write_html_s2_div1_begin(site)
write_html_s2_div1(site)
date_heure()
write_html_table(site, colonne1, colonne2)
write_html_s2_div1_end(site)
write_html_s2_end(site)

write_html_s3_begin(site)
write_html_s3_div2_begin(site)
write_html_s3_div2(site)
prevision_s()
write_html_s3_div2_end(site)
write_html_s3_end(site)
write_html_body_end(site)
write_html_end(site)

site.close()
