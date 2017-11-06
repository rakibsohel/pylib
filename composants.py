#!/usr/bin/env python
# -*- coding: utf-8 -*-

import re
import win32com.client
from   datetime import date
import os
import shutil


class wd:

    def __init__(self, NomDoc):        

        #Variables
        self._debug = False
        self._nbeTab = 0
        self._lst_nomCon = []
        
        #Variable nom du document word.
        self.NomDoc = NomDoc

        #Constante de séparation.
        self.WORD_END_CELL_SEPARATOR = u'\r\x07'

        #Lancement de l'application word.
        self.wd = win32com.client.Dispatch('Word.Application');
        self.wd.Visible = False

        #Objet word.
        self.doc = self.wd.Documents.Add(self.NomDoc)

        #Maj var
        self._nbeTab = self.doc.Tables.Count

        #Maj Tab
        self.update()

    def _get_debug(self):
        return(self._debug)
        
    def _set_debug(self, val_debug):
        self._debug = val_debug
    
    debug = property(_get_debug,_set_debug)        

    def _get_nbeTab(self):
        return(self._nbeTab)
    
    nbeTab = property(_get_nbeTab)

    def _get_lst_nomCon(self):
        return(self._lst_nomCon)
    
    lst_nomCon = property(_get_lst_nomCon)

    def _get_lst_AllCon(self):
        return(self._lst_AllCon)
    
    lst_AllCon = property(_get_lst_AllCon)     

    def typeTab(self,numTab):

        #Variables
        self.numTab = numTab
        self.bOk    = True
        
        #Informations table
        self.tab = self.doc.Tables[self.index]
        self.nbeLigne = self.doc.Tables[self.index].Rows.Count

        if((str(self.tab.Cell(1,1)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper()) != "CONNECTEUR"):
            self.bOk = False

        return(self.bOk)

    def nomConnecteur(self,numTab):

        #Variables
        self.numTab = numTab
        
        #Informations table
        self.tab      = self.doc.Tables[self.index]
        self.nbeLigne = self.doc.Tables[self.index].Rows.Count

        #self.lst_nomCon.append((str(self.tab.Cell(1,2)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper()))

    def typeConnecteur(self,numTab):        

        #Variables
        self.numTab = numTab
        self.iFcon  = False
        
        #Informations table
        self.tab = self.doc.Tables[self.index]
        self.nbeLigne = self.doc.Tables[self.index].Rows.Count

        if((str(self.tab.Cell(1,3)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper()) == "CONNECTEUR"):
           self.iFcon  = False

        return(self.iFcon)

    def get(self,tableau):
        
        #Variables
        self.index    = tableau
        self.chaine   = ""
        self.lst_getTab = []

        #Informations table
        self.tab = self.doc.Tables[self.index]
        self.nbeLigne = self.doc.Tables[self.index].Rows.Count

        if(self.tab is not None):
            if((str(self.tab.Cell(1,1)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper()) == "CONNECTEUR"):

                #Vérification de la conformité des colonnes
                if( (str(self.tab.Cell(2,1)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper() == "PIN") and
                    (str(self.tab.Cell(2,2)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper() == "LABEL") and
                    (str(self.tab.Cell(2,3)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper() == "EQUIPOTENTIELLE") and
                    (str(self.tab.Cell(1,3)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper() != "")):

                    self.nom = ""
                    self.nom = str(self.tab.Cell(1,2)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper()

                    if(self._lst_nomCon.count(self.nom) < 1):
                        self._lst_nomCon.append(self.nom)
                  
                    #Parcours Tableau
                    self.ligne = 3
                    while (self.ligne != (self.nbeLigne+1)) :
                        self.chaine = (str(self.tab.Cell(1,3)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper()         +"|"+
                                       str(self.tab.Cell(1,2)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper().upper() +"|"+
                                       str(self.tab.Cell(self.ligne,1)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper()+"|"+                               
                                       str(self.tab.Cell(self.ligne,2)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper()+"|"+
                                       str(self.tab.Cell(self.ligne,3)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper()+"|"                             
                                       )
                        if(len(str(self.tab.Cell(self.ligne,4)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper()) < 1):
                            self.chaine = self.chaine +"N.U"
                        else:
                            self.chaine = self.chaine + str(self.tab.Cell(self.ligne,4)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper()
                    
                        self.lst_getTab.append(self.chaine)
    
                        if(self._debug == True):
                            print(self.chaine)
                    
                        self.ligne = self.ligne + 1

                    return(self.lst_getTab)
                else:
                    print("Erreur conformité du tableau -> "+str(self.index)+" : "+str(self.tab.Cell(1,2)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper().upper())        

    def update(self):

        self._lst_AllCon = []
        self.i = 0

        while self.i != self.nbeTab:
            if(self.get(self.i) is not None):
                self.lst_AllCon.append(self.get(self.i))            
            self.i = self.i + 1

        return(self._lst_AllCon)            

    def quit(self):
        self.wd.Quit() 

class library(wd):

    #Constructeur objet
    def __init__(self,pathDoc):

        #Variable document input ".doc".
        self.pathDoc = pathDoc

        #Variables document export ".bom".
        self._pathBom = "C:/Python34/Scripts/Doc2Lst/out"        

        #Chemin de la librairie.
        self._pathLib = "C:/AFT/AFT/BIBLIO"

        #Constructeur classe mère.
        wd.__init__(self,self.pathDoc)


    #Mutateur et Accesseur#
    def _get_pathLib(self):
        return(self._pathLib)

    def _set_pathLib(self,val_pathLib):
        self._pathLib = val_pathLib

    pathLib = property(_get_pathLib,_set_pathLib)

    def _get_pathBom(self):
        return(self._pathBom)

    def _set_pathBom(self,val_pathBom):
        self._pathBom = val_pathBom

    pathBom = property(_get_pathBom,_set_pathBom)    

    #Cette fonction créer un fichier ".bom".
    def toBom(self,export):

        #Créationet/ou ouverture du fichier d'export.
        self.bom = open(self._pathBom+export+".bom","w")

        #Récupération de la liste des les lignes des tableaux.
        self.lst_Ligne = self.lst_AllCon

        #Récupération de la liste des connecteurs.
        self.lst_Con   = self.lst_nomCon
        print(self.lst_Con)
        #Parcours de la liste des connecteurs.#
        for self.con in self.lst_Con:

            #Variables#

            #Nombre de pin du connecteur.
            self.nbePin  = 0

            #Liste temporaire des points du connecteur
            self.lst_tmp =[]

            #Variable contenant le type de connecteur.
            self.typeCon = "N.U"

            #Parcours des lignes des connecteurs des tableaux#
            #Tableaux deux dimensions.#
            for self.l in self.lst_Ligne:

                #print(self.l)
                if(self.l is not None):
                    #print(self.l)
                    #Parcours des lignes du connecteur.#
                    for self.el in self.l:

                        #Division de la ligne#
                        self.tabLigne = self.el.split("|")

                        #S'il y a le nom du connecteur courant dans la ligne#
                        if(self.con == self.tabLigne[1]):

                            #Ajout de la ligne dans la liste temporaire <Format bom>.
                            self.lst_tmp.append(self.tabLigne[2]+"|"+self.tabLigne[3]+"|")

                            #Maj de la variable du type de connecteur.
                            self.typeCon = self.tabLigne[0]

                            #Incrémentation de la ligne et comptage du nombre de pin.
                            self.nbePin = self.nbePin + 1
                        
            if(self.checkLib(self.typeCon) == self.nbePin):
                
                #Ecriture de l'entête dans le fichier ".bom".
                self.bom.write("#|"+self.typeCon+"|"+self.con+"||\n")
                self.bom.write(">|"+self.typeCon+"|"+str(self.nbePin)+"| |RESSOURCE|\n")

                #Ecriture des lignes du connecteur#            
                for self.n in self.lst_tmp:

                    #Ecriture de la ligne.
                    self.bom.write(self.n+"\n")
            else:
                print("ERREUR LIB -> "+self.typeCon+"|"+self.con)                
                self.bom.close()
                os.remove(self._pathBom+export+".bom")
                break

            #Suppression de la lste temporaire.
            del self.lst_tmp

        #Femreture du fichier ".bom".
        self.bom.close()

    #Cette fonction permet de vérifer si le composant existe dans le fichier biblio.txt#
    def checkLib(self,composant):

        #Variable de vérification.
        self.bFound     = 0

        #Nom du composant
        self.composant = composant

        #Si le composant existe dans le fichier Biblio et un fichier du composant existe aussi.#
        if((os.path.isfile(self._pathLib+"/Biblio.txt") == True) and
           (os.path.isfile(self._pathLib+"/"+self.composant+".txt") == True)):
           
           self.ficLib = open(self._pathLib+"/Biblio.txt","r")
           self.content = self.ficLib.readlines()

           for self.el in self.content:
               if (re.search(self.composant,self.el)) is not None:
                   self.ficCmp     = open(self._pathLib+"/"+self.composant+".txt")
                   self.contentCmp = self.ficCmp.readlines()

                   #Compte le nombre de pin du composant à partir de la bibliothèque.#
                   self.bFound     = 0
                   for self.lCmp in self.contentCmp:
                       if(self.lCmp.strip() != ""):
                           #print(self.lCmp)
                           self.bFound = self.bFound + 1
                   
                   self.ficCmp.close()

           self.ficLib.close()
           
        else:
            self.bFound = 0            

        #Renvoie le nombre de pin.
        #print(self.bFound-1)
        return(self.bFound-1)

    #Fermeture de l'objet word.#
    def quit(self):        
        wd.quit(self)
        

def main(sel):

    if(sel.upper() == "BOM"):
        bm = library("C:/Python34/Scripts/Doc2Lst/travail/INTERFACE_DE_TEST_CUC05_1A.docx")
        bm.pathBom = "C:/Python34/Scripts/Doc2Lst/out"
        bm.toBom("/INTERFACE_DE_TEST_CUC05_2A")
        bm.quit()
        
    if(sel.upper() == "DOC"):
        #Création de l'objet.
        FileWord = wd("C:/Python34/Scripts/Doc2Lst/travail/INTERFACE_DE_TEST_CUC05_1A.docx")    
    
        lst = FileWord.lst_nomCon
        print(lst)                   
    
        FileWord.quit()

if __name__=="__main__":
    main("DOC")
    #main("BOM")
