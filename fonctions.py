#Author Rivo Lalaina RAJAONARIVONY
#coding utf _8
#Uu programme qui génerera les tâches en année préparatoire.



import time , datetime, os ,xlsxwriter
import tkinter.messagebox as tkmsg 
from tkinter import *
from tkinter.font import *

alphabet= "ABCDEFGHIJKKLMNOPQRSTUVWXYZ"



class grand_menage():
#une class qui affecte des valeurs a des variables dont le but est  de pouvoir modifier facilement depuis l interface.

	def menage(self):
		self.grandM1 = 'Frigo,aspirateur, fours à gaz ,cuisinière'
		self.grandM2 = 'Tables,leviers,carrelage et chariots à la cuisine' 
		self.grandM3 = 'Vitre intérieur et mur de la cuisine'
		self.grandM4 = 'Intendance et laver les chiffons'
		self.grandM5 = "Locaux ordures ,dehors cuisines et vitres exterieure"
		self.grandM6 = "Tables,chaises,toiles d'arraignés,carrelage et vitre intérieure du réfectoire"
		#  leurs valeurs par defaut
class tache_name_changer():
	""" une class qui affecte une liste de valeur au variable taches dont le but est 
de pouvoir modifier facilment depuis l Interface """
	def  t_changer(self):
		self.taches = ['Cuisine', 0, 'Couvert&Refectoire', 0, 'Vaisselle']
		#  valeur par defaut de taches

ind = tache_name_changer() #  cree une instance pour la class
gm = grand_menage() #  crée une instance pour cetteclass
ind.t_changer()
 


'''une classe qui convertira le mois en français.'''
class mois_fr():
    
    def month(self,jour):
        mois ={ "01": "Janvier", "02":"Fevrier", "03":"Mars", "04":"Avril", "05":"Mai", "06":"Juin", "07":"Juillet", "08":"Août", "09":"Septembre", "10":"Octobre", "11":"Novembre", "12":"Decembre"}

        for x,y in mois.items ():
            if str(datetime.date.strftime(datetime.date.fromordinal (jour), "%m")) ==x:
                return y #La difference entre return et print et que la fonction return appartient toujours dans une fonction
        # x= les cles chiffres, y le mois corespondant.


'''La classe qui génerera notre fichier excel.'''
class ecrire():
    def __init__(self):
        
        # Cette fonction aura pour but d'initier notre emploi du temps le lundi.
        if time.strftime('%a') == 'Mon':
            self.jour =int(datetime.date.toordinal(datetime.date.today()) ) -1

        elif time.strftime('%a') == 'Tue':
            self.jour = int( datetime.date.toordinal(datetime.date.today()) )  -2

        elif time.strftime('%a') == 'Wed':
            self.jour = int( datetime.date.toordinal(datetime.date.today()) ) - 3

        elif time.strftime('%a') == 'Thu':
            self.jour = int( datetime.date.toordinal(datetime.date.today()) ) - 4

        elif time.strftime('%a') == 'Fri':
            self.jour = int( datetime.date.toordinal(datetime.date.today()) ) - 5

        elif time.strftime('%a') == 'Sat':
            self.jour = int( datetime.date.toordinal(datetime.date.today()) ) - 6

        else : 
            self.jour = int( datetime.date.toordinal(datetime.date.today()) )
                        
        self.datename = str(datetime.date.strftime(datetime.date.fromordinal(self.jour), "%d-%m-%Y"))
        self.datefin = str(datetime.date.strftime(datetime.date.fromordinal(self.jour+6), "%d-%m-%Y"))
        mon = mois_fr()
        '''  Une fonction qui va créer notre ficher excel.'''		
    
    def initial(self):        
    
        try :
            self.fichier =xlsxwriter.Workbook ( r"fichier_excel\\Tache"+str(self.datename)+ ".xlsx" )
                       
        except  FileNotFoundError:
            os.system("md fichier_excel")
            tkmsg.showinfo ("Dossier", "Un  dossier a été créer.")
            self.fichier =xlsxwriter.Workbook ( r"fichier_excel\\Tache"+str (self.datename)+ ".xlsx" )
        except PermissionError:
            self.fichier = xlsxwriter.Workbook(r"fichier_excel\\Tache"+str(self.datename)+"_1.xlsx")
            tkmsg.showinfo ("Dossier", "Un  dossier a été créer.")

        self.feuille_excel =self.fichier.add_worksheet(self.datename + '-' +self.datefin)

        self.taches = ['Cuisine', 0, 'Couvert&Refectoire', 0, 'Vaisselle']
        self.mmm =["Matin", "Midi", "Soir"]
        self.frat = list("FRAT {}".format(x+1) for x in range (6))
        self.cm_s = ["APS1", "APS2"]
        self.cm_l = ["APL1" , "APL2"]

        self.feuille_excel.set_landscape()
        self.feuille_excel.hide_gridlines(2)  

    '''Une fonction qui se chargera de créer la tâche de la semaine d'après.'''
    def initial_next(self):
        self.datename = str(datetime.date.strftime(datetime.date.fromordinal(self.jour + (7*self.nextchanger)), "%d-%m-%Y"))

        try :
            self.fichier =xlsxwriter.Workbook ( r"fichier_excel\\Tache"+str(self.datename)+ ".xlsx" )
                       
        except  FileNotFoundError:
            os.system("mkdir fichier_excel")
            tkmsg.showinfo ("Dossier", "Un  dossier a été créer.")
            self.fichier =xlsxwriter.Workbook ( r"fichier_excel\\Tache"+str (self.datename)+ ".xlsx" )
        except PermissionError:
            self.fichier = xlsxwriter.Workbook(r"fichier_excel\\Tache"+str(self.datename)+"_1.xlsx")
            tkmsg.showinfo ("Dossier", "Un  dossier a été créer.")



        
        self.feuille_excel =self.fichier.add_worksheet(self.datename + '-' +self.datefin)
        self.taches = ['Cuisine', 0, 'Couvert&Refectoire', 0, 'Vaisselle']
        self.mmm =["Matin", "Midi", "Soir"]
        self.frat = list("FRAT {}".format(x+1) for x in range (6))
        self.cm_s = ["APS1", "APS2"]
        self.cm_l = ["APL1" , "APL2"]
        self.feuille_excel.set_landscape()
        self.feuille_excel.hide_gridlines(2) 
    '''La fonction qui écrirera la date.'''

    def forme (self) :

        self.merge_format_pT = self.fichier.add_format({'align': 'center', 'bold': 1, 'fg_color': 'gray','border':1,'italic':True})
        self.fin_format = self.fichier.add_format({'align': 'center', 'bold': 1, 'font_size':14})
        self.centrer = self.fichier.add_format({'align': 'center','valign': 'vcenter', 'border': 1,'bold': 2,'bg_color': 'gray'})        
        self.cuisine_format = self.fichier.add_format({'align': 'center','valign': 'vcenter', 'font_size': 11, 'border':1})
        self.mmm_format = self.fichier.add_format({'align': 'center','bg_color': 'gray', 'border':1}) 
        self.ajout_format=self.fichier.add_format({'font_size': 11, 'border':1,'align':'left','valign':'top'})
        self.merge_format_TV = self.fichier.add_format({'align': 'center', 'bold': 1, 'valign': 'vcenter', 'border':1})
    
    def ecrire_date(self):
        self.i = 0
        mon = mois_fr()

        for x in range (0, 12, 3):

            if str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Sun":
                self.days = "Dimanche"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Mon":
                self.days = "Lundi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Tue":
                self.days = "Mardi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Wed":
                self.days = "Mercredi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Thu":
                self.days = "Jeudi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Fri":
                self.days = "Vendredi"

            else:
                self.days = "Samedi"
        
            self.date = str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%d {} %Y".format(mon.month(self.jour + self.i))))
            self.feuille_excel.merge_range('{}3:{}3'.format(alphabet[x+2], alphabet[x+4]), "{} {}".format(self.days, self.date), self.merge_format_TV)
            self.i += 1 

        for x in range(0, 9, 3):

            if str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Mon":
                self.days = "Lundi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Tue":
                self.days = "Mardi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Wed":
                self.days = "Mercredi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Thu":
                self.days = "Jeudi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Fri":
                self.days = "Vendredi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Sat":
                self.days = "Samedi"

            else:
                self.days = "Dimanche"
		
            self.date = str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%d {} %Y".format(mon.month(self.jour + self.i))))
            self.feuille_excel.merge_range('{}12:{}12'.format(alphabet[x+2], alphabet[x+4]), "{} {}".format(self.days, self.date), self.merge_format_TV)
            self.i += 1
			
        del x
        del self.i


    def ecrire_next_date(self):
        self.jour = self.jour + (7 * self.nextchanger)
        self.i = 0
        mon = mois_fr()

        for x in range(0, 12, 3):

            if str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Mon":
                self.days = "Lundi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Tue":
                self.days = "Mardi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Wed":
                self.days = "Mercredi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Thu":
                self.days = "Jeudi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Fri":
                self.days = "Vendredi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Sat":
                self.days = "Samedi"

            else:
                self.days = "Dimanche"

            self.date = str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%d {} %Y".format(mon.month(self.jour + self.i))))
            self.feuille_excel.merge_range('{}3:{}3'.format(alphabet[x+2], alphabet[x+4]), "{} {}".format(self.days, self.date), self.merge_format_TV)
            self.i += 1


        for x in range(0, 9, 3):

            if str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Mon":
                self.days = "Lundi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Tue":
                self.days = "Mardi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Wed":
                self.days = "Mercredi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Thu":
                self.days = "Jeudi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Fri":
                self.days = "Vendredi"

            elif str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%a")) == "Sat":
                self.days = "Samedi"

            else:
                self.days = "Dimanche"

            self.date = str(datetime.date.strftime(datetime.date.fromordinal(self.jour + self.i), "%d {} %Y".format(mon.month(self.jour + self.i))))
            self.feuille_excel.merge_range('{}12:{}12'.format(alphabet[x+2], alphabet[x+4]), "{} {}".format(self.days, self.date), self.merge_format_TV)
            self.i += 1 
                                                       
        self.jour = self.jour + (7 * self.nextchanger)
        self.i = 0
        mon = mois_fr()


        del x
        

    
    def text_principal_next (self):
        
        #On instancie l'objet mon pour notre mois français.
        mon = mois_fr()
        self.date_start = str(datetime.date.strftime(datetime.date.fromordinal(self.jour + 7 * self.nextchanger), "%d {} %Y".format(mon.month(self.jour + 7 * self.nextchanger))))
        self.feuille_excel.merge_range('D1:K1', "Repartition des taches à l'Atrium pour la semaine du {})".format(self.date_start), self.merge_format_pT)

        #L'intérêt de ce mot clé del est de pouvoir les modifier pour les prochains fichiers.
        del self.date_start


        '''La fonction qui écrira les tâches.'''
    
    def select_date(self):
        self.date0 = datetime.date(2020, 2, 9) # La date supposée comme la date du debut, l'année 0 :-p
        self.date_x = datetime.date.today()  # la variable contenant la date actuel

        if self.date0 == self.date_x:
            return 0
        else:
            return int(str(self.date_x - self.date0).split(" ")[0])  # Le nombre de jour entre la date actuel et la date initial
    
    def sesame_logo(self):                                                     
        try:
            self.feuille_excel.insert_image("M20", "Images\\sesame.png")
        except:
            pass
        
    def select_week(self):
        self.temps = self.select_date()
        return self.temps // 7
    def ecrire_tache(self):

        self.feuille_excel.set_column(0, 0, 18)         
         # Dans le colonne A depui le A5 à A10 en incrémentant de2."
        self.feuille_excel.merge_range('A3:A4', 'Taches', self.merge_format_TV)        
        for x in range(5, 11, 2):
            self.feuille_excel.merge_range('A{}:A{}'.format(x,x+1), ind.taches[x-5],self.centrer)

        self.feuille_excel.merge_range ('A12:A13', 'Taches',self.merge_format_TV) #self.merge_format_pT)

        for x in range (14,20,2): # Dans le colonne A depui le A5 à A10 en incrémentant de2."
            self.feuille_excel.merge_range('A{}:A{}'.format(x,x+1), ind.taches[x-14],self.centrer)
        del  x #On supprime X pour pouvoir l'utiliser après comme une variable.
        #self.fichier.close()

        #La fonction qui écrira la liste .
    
    def ecrire_partie(self):


        self.feuille_excel.set_column(1, 1, 26)     # elargir le colonne
        self.feuille_excel.set_column(3, 3, 12)     # elargir le colonne
        self.feuille_excel.set_column(6, 6, 12)     # elargir le colonne
        self.feuille_excel.set_column(9, 9, 15)     # elargir le colonne  
        self.feuille_excel.set_column(12, 12, 12) 
        self.feuille_excel.set_row(9,30) #elargir la hauteur de 
        self.feuille_excel.set_row(18,30) #elargir la hauteur de la colonne

        self.feuille_excel.merge_range('B3:B4','Partie', self.centrer)
        self.feuille_excel.write('B5', 'cuisine', self.cuisine_format)
        self.feuille_excel.write('B6', 'wc-Int', self.cuisine_format)
        self.feuille_excel.write('B7', 'Couvert (1er-2ème vague)', self.cuisine_format)                                                       
        self.feuille_excel.write('B8','Refectoire (1er_2ème vague)' , self.cuisine_format)                        
        self.feuille_excel.write('B9','Vaisselle' , self.cuisine_format)
        self.feuille_excel.write('B10', 'Rice cooker \n Externe de la cuisine',self.ajout_format)
		

        self.feuille_excel.merge_range('B12:B13', 'Partie', self.centrer)
        self.feuille_excel.write('B14', 'cuisine', self.cuisine_format)
        self.feuille_excel.write('B15', 'wc-Int', self.cuisine_format)
        self.feuille_excel.write('B16', 'Couvert (1er-2ème vague)', self.cuisine_format)                                                       
        self.feuille_excel.write('B17', 'Refectoire (1er_2ème vague)' , self.cuisine_format)                        
        self.feuille_excel.write('B18', 'Vaisselle', self.cuisine_format)
        self.feuille_excel.write('B19', 'Rice cooker \n Externe de la cuisine', self.ajout_format)                                                               

    def ecrire_mmm(self):

        self.feuille_excel.write("L4", 'Matin', self.mmm_format)
        self.feuille_excel.write("M4", 'Midi', self.mmm_format)
        self.feuille_excel.write("N4", 'Soir', self.mmm_format)                
        for x in range(9):
            self.feuille_excel.write('{}4'.format(alphabet[x+2]), self.mmm[x%3], self.mmm_format)
        for x in range(9):
            self.feuille_excel.write('{}13'.format(alphabet[x+2]), self.mmm[x%3], self.mmm_format)
        del x

    def ecrire_grand_menage(self):
                                                      
        self.feuille_excel.merge_range('L12:N12', 'Samedi de 16h à 17h', self.merge_format_pT)
        self.feuille_excel.merge_range('L13:N13', 'Grand Menages', self.cuisine_format)
        gm=grand_menage()
        gm.menage()

        for x in range(14, 20):
            self.feuille_excel.merge_range('L{}:M{}'.format(x, x), 'GM Tache {}'.format(x - 13), self.mmm_format)

        self.feuille_excel.write("A21", "GM Tache 1:  '{}' ".format(gm.grandM1))
        self.feuille_excel.write("A22", "GM tache 2:  '{}' ".format(gm.grandM2))
        self.feuille_excel.write("A23", "GM Tache 3:  '{}' ".format(gm.grandM3))
        self.feuille_excel.write("A24", "GM Tache 4:  '{}' ".format(gm.grandM4))
        self.feuille_excel.write("A25", "GM Tache 5:  '{}' ".format(gm.grandM5))
        self.feuille_excel.write("A26", "GM Tache 6:  '{}' ".format(gm.grandM6))

    def ecrire_gmvariable(self):
                                                       
        self.week = self.select_week()
        self.week = self.week + self.nextchanger
        y = 0
		
        for x in range((7*self.week), ((7*self.week) + 6)):
            self.feuille_excel.write("N{}".format(y + 14), self.frat[x % 6], self.cuisine_format)
            y += 1
        del x
        #self.fichier.close()
    
    def ecrire_variable(self):
                                                       
        self.frat= ["frat 5", "frat 6", "frat 3", "frat 4", "frat 1", "frat 2","frat 6", "frat 5", "frat 4", "frat 3", "frat 2", "frat 1"]
        self.frat_inverse = ["frat 6", "frat 5", "frat 4", "frat 3", "frat 2", "frat 1", "frat 5", "frat 6", "frat 3", "frat 4", "frat 1", "frat 2",]
        self.week = self.select_week()
        self.week = self.week + self.nextchanger
  
        y = 0
        self.week = 2 * self.week

        for x in range((7*self.week), ((7*self.week) + 6)):

            self.feuille_excel.write("C{}".format(y + 5), self.frat[x % 12], self.cuisine_format)
            self.feuille_excel.write("E{}".format(y + 5), self.frat_inverse[x % 12], self.cuisine_format)
            self.feuille_excel.write("F{}".format(y + 5), self.frat_inverse[(x +2) % 12], self.cuisine_format)
            self.feuille_excel.write("H{}".format(y + 5), self.frat[(x + 2) % 12], self.cuisine_format)
            self.feuille_excel.write("I{}".format(y + 5), self.frat[(x + 4) % 12], self.cuisine_format)
            self.feuille_excel.write("K{}".format(y + 5), self.frat_inverse[(x + 4) % 12], self.cuisine_format)
            self.feuille_excel.write("L{}".format(y + 5), self.frat_inverse[x  % 12], self.cuisine_format)
            self.feuille_excel.write("N{}".format(y + 5), self.frat[x % 12], self.cuisine_format)

            self.feuille_excel.write("C{}".format(y + 14), self.frat[(x+2) % 12], self.cuisine_format)
            self.feuille_excel.write("E{}".format(y + 14), self.frat_inverse[(x+2) % 12], self.cuisine_format)
            self.feuille_excel.write("F{}".format(y + 14), self.frat_inverse[(x + 4) % 12], self.cuisine_format)
            self.feuille_excel.write("H{}".format(y + 14), self.frat[(x + 4) % 12], self.cuisine_format)
            self.feuille_excel.write("I{}".format(y + 14), self.frat[x  % 12], self.cuisine_format)
            self.feuille_excel.write("K{}".format(y + 14), self.frat_inverse[x % 12], self.cuisine_format)
            
            y += 1
        del x
        del y

    def ecrire_varmidi(self):
        self.week = self.select_week()
        self.week = self.week + self.nextchanger
        y = 0
        for x in range((7*self.week), ((7*self.week) + 6)):
            self.feuille_excel.write("D{}".format(y + 5), self.frat[(x + 2) % 6], self.cuisine_format)
            self.feuille_excel.write("J{}".format(y + 14), self.frat[(x + 4) % 6],self.cuisine_format)
            y += 1

        for n in range(6, 11, 3):
            self.feuille_excel.write("{}5".format(alphabet[n]), "Cuisiniers", self.cuisine_format)
            self.feuille_excel.write("{}9".format(alphabet[n]), "Cuisiniers", self.cuisine_format)
            self.feuille_excel.write("{}7".format(alphabet[n]), "Cuisiniers", self.cuisine_format)            
            self.feuille_excel.write("{}10".format(alphabet[n]), "Cuisiniers", self.cuisine_format)


			
        self.autrelist = ["G","J","M","D","G2","D2","M2"]
		
        for autrelist in self.autrelist:

            if  autrelist =="G" or autrelist =="J" or autrelist =="M"  :               
                self.feuille_excel.write("{}7".format(autrelist),'APS1-APL1', self.cuisine_format)	
                self.feuille_excel.write("{}8".format(autrelist),'APS2-APL2', self.cuisine_format)
                self.feuille_excel.write("{}5".format(autrelist),"Cuisiniers", self.cuisine_format)	
                self.feuille_excel.write("{}9".format(autrelist),"Cuisiniers", self.cuisine_format)					
               
            elif autrelist == "D2" :
                self.feuille_excel.write("D14", "Cuisiniers", self.cuisine_format)	
                self.feuille_excel.write("D18", "Cuisiniers", self.cuisine_format)

            else:
                self.feuille_excel.write("D16", 'APS1-APL1', self.cuisine_format)	
                self.feuille_excel.write("D17", 'APS2-APL2', self.cuisine_format)
        del y
        del autrelist
        del self.autrelist
        del n
        self.autreList = ["G"]
        for autreList in self.autreList:
            self.feuille_excel.write("{}14".format(autreList),"Cuisiniers", self.cuisine_format)	
            self.feuille_excel.write("{}18".format(autreList),"Cuisiniers", self.cuisine_format)
            self.feuille_excel.write("{}16".format(autreList),'APS1-APL1', self.cuisine_format)	
            self.feuille_excel.write("{}17".format(autreList),'APS2-APL2', self.cuisine_format)            	
        del autreList
        del self.autreList                        
    def __fin__(self):
        self.feuille_excel.merge_range('C30:K30',"Ce n'est qu'en changeant l'éducation que l'on peut changer le monde." , self.merge_format_TV) 
        tkmsg.showinfo("Remarque","Votre tâche est enregistrer dans le fichier 'fichier_excel'")
        self.feuille_excel.write('D6',None,self.cuisine_format)
        self.feuille_excel.write('J15',None,self.cuisine_format)   
        self.feuille_excel.write('D19',None,self.cuisine_format)
        self.feuille_excel.write('G19',None,self.cuisine_format)
        self.feuille_excel.write('M10',None,self.cuisine_format)            
        self.fichier.close()


class interface():
    def __init__(self):
        self.root = Tk()
        self.root.title("Tâche en année préparatoire")
        self.root.geometry("1366x760")
        self.root.iconbitmap(r'Images\\sesame.ico')
        self.nombreErreur = 0	
    def get_start(self):

        excel = ecrire()

        excel.nextchanger =0
        excel.initial()
        excel.forme()
        excel.ecrire_date()

        excel.text_principal_next()
        excel.ecrire_partie()
        excel.select_date()
        excel.ecrire_mmm()
        excel.ecrire_tache()
        excel.ecrire_grand_menage()
        excel.sesame_logo()
        excel.select_week()
        excel.ecrire_gmvariable()
        excel.ecrire_variable()
        excel.ecrire_varmidi()
        excel.__fin__()

        if self.nombreErreur < 2:
                self.nombreErreur += 1

        elif self.nombreErreur == 2:
            tkmsg.showerror("Erreur n° {}".format(self.nombreErreur), "Verifier s'il y pas de fichier Tache excel ouvert, si oui Fermer")
            self.nombreErreur += 1

        elif self.nombreErreur > 2 and self.nombreErreur < 5:
            tkmsg.showerror("Erreur n° {}".format(self.nombreErreur), "Verifier si un dossier 'Output' est présent dans le dossier contenant l'excecutable, sinon créé")
            self.nombreErreur += 1
        else:
            tkmsg.showerror("Erreur", 'Veuiller Contacter le Developper pour resoudre le probleme')
			
			
    def get_next(self):
        
        try:
            
            excel = ecrire()
            excel.nextchanger = 1
            excel.initial_next()
            excel.forme()
            excel.ecrire_next_date()
            excel.text_principal_next()
            excel.ecrire_partie()
            excel.select_date()
            excel.ecrire_mmm()
            excel.ecrire_tache()
            excel.ecrire_grand_menage()
            excel.sesame_logo()
            excel.select_week()
            excel.ecrire_gmvariable()
            excel.ecrire_variable()
            excel.ecrire_varmidi()
            excel.__fin__()

		
        except:              
            if self.nombreErreur < 2:
                self.nombreErreur += 1

            elif self.nombreErreur == 2:
                tkmsg.showerror("Erreur n° {}".format(self.nombreErreur), "Verifier s'il y pas de fichier Tache excel ouvert, si oui Fermer")
                self.nombreErreur += 1

            elif self.nombreErreur > 2 and self.nombreErreur < 5:
                tkmsg.showerror("Erreur n° {}".format(self.nombreErreur), "Verifier si un dossier 'Output' est présent dans le dossier contenant l'excecutable, sinon créé")
                self.nombreErreur += 1

            else:
                tkmsg.showerror("Erreur", 'Veuiller Contacter le Developper pour resoudre le probleme')


    def get_next_next(self):

        try: 
            test = int(self.entry_champ.get())
            test = True

        except ValueError:
            test = False

        if test:
            try:
                excel = ecrire()
                excel.nextchanger = int(self.entry_champ.get())
                excel.initial_next()
                excel.forme()
                excel.ecrire_next_date()

                excel.text_principal_next()
                excel.ecrire_partie()
                excel.select_date()
                excel.ecrire_mmm()
                excel.ecrire_tache()
                excel.ecrire_grand_menage()
                excel.sesame_logo()
                excel.select_week()
                excel.ecrire_gmvariable()
                excel.ecrire_variable()
                excel.ecrire_varmidi()
                excel.__fin__()

                
            except:            
                if self.nombreErreur < 2:
                    self.nombreErreur += 1

                elif self.nombreErreur == 2:
                    tkmsg.showerror("Erreur n° {}".format(self.nombreErreur), "Verifier s'il y pas de fichier Tache excel ouvert, si oui Fermer")
                    self.nombreErreur += 1

                elif self.nombreErreur > 2 and self.nombreErreur < 5:
                    tkmsg.showerror("Erreur n° {}".format(self.nombreErreur), "Verifier si un dossier 'Output' est présent dans le dossier contenant l'excecutable, sinon créé")
                    self.nombreErreur += 1

                else:
                    tkmsg.showerror("Erreur", 'Veuiller Contacter le Developper pour resoudre le probleme')
        else:
            tkmsg.showerror("Erreur d'entrée","Il semblerait que vous avez entrer un type non valable \n L'entrée doit être un entier")

        del test
    def button(self):
        self.button_start_font = Font(family ="Times New Rowan", size = -20, weight = "bold")
        self.button_start = Button(self.root, text = "Créer la tâche de cette semaine", font = self.button_start_font, bg = "gray", height = 3, width = 40, fg = 'green',  command = self.get_start)
        self.button_start.place(x ="425", y = "425")

        self.button_start_font = Font(family ="Times New Rowan", size = -20, weight = "bold")
        self.button_start = Button(self.root, text = "Créer la tâche de la semaine prochaine", font = self.button_start_font, bg = "gray", height = 3, width = 49, fg = 'green',  command = self.get_next)
        self.button_start.place(x ="385", y = "525")

    def label_bord(self):
        self.font_text1 = Font(family = "Arial", size = -12, weight = "bold", underline = True)
        self.text1 = Label(self.root, text = "Avancer de plusieurs semaines?", bg = "white", font = self.font_text1)
        self.text1.place(x = "1128", y = "70")

        self.entry_variable = IntVar()
        self.entry_champ = Entry(self.root, textvariable = self.entry_variable, width = 5)
        self.entry_champ.place(x="1230", y ="100")
        self.entry_text_s = Label(self.root,font = Font(family = 'Arial', size = -14, weight = 'bold'), bg = "#e8e8c8", text = "Avancer de ").place(x="1135", y ="100")
        self.entry_button = Button(self.root, text= "GENERER",font = Font(family = 'Arial', size = -14, weight = 'bold'), command = self.get_next_next ).place(x="1130", y="125")
        self.entry_text_e = Label(self.root, font = Font(family = 'Arial', size = -14, weight = 'bold'), bg = "#e8e8c8", text = "semaines").place(x="1265", y ="100")
        del self.font_text1
    def text_final(self):
        self.font_text1 = Font(family = "Arial", size = -12, underline = False)
        self.final_text=Label(self.root ,text="@copyright 2002",font=self.font_text1,bg ="#e8e8c8").place (x="570",y="650")
    def __final__(self):
        self.root.mainloop()

rivo=interface	()
rivo.button	()
rivo.label_bord()
rivo.text_final()
rivo.__final__()


