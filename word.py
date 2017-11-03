#!/usr/bin/env python
# -*- coding: utf-8 -*-

#******************************************************************
# Programme   : WordToBom
# Auteur      : Bepari Rakib
# Date        : 02/11/17
# Indice      : A
# Description : Ce script génère un fichier .bom à partir d'un DCM.
#******************************************************************

#Librairie
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

        self.lst_nomCon.append((str(self.tab.Cell(1,2)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper()))

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

        if((str(self.tab.Cell(1,1)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper()) == "CONNECTEUR"):

            #Vérification de la conformité des colonnes
            if( (str(self.tab.Cell(2,1)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper() == "PIN") and
                (str(self.tab.Cell(2,2)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper() == "LABEL") and
                (str(self.tab.Cell(2,3)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper() == "EQUIPOTENTIELLE") and
                (str(self.tab.Cell(1,3)).rstrip(self.WORD_END_CELL_SEPARATOR).rstrip().lstrip().upper() != "")):
    
                self.nomConnecteur(self.index)
                  
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
            self.lst_AllCon.append(self.get(self.i))            
            self.i = self.i + 1

        return(self._lst_AllCon)            

    def quit(self):
        self.wd.Quit()    


def main():

    #Création de l'objet.
    FileWord = wd("C:/Python34/Scripts/Doc2Lst/travail/INTERFACE_DE_TEST_CUC05_1A.docx")    
    
    lst = FileWord.lst_AllCon
    print(len(lst))
    
    for l in lst:
        for j in l:
            print(j)
    
    FileWord.quit()
    
if __name__ == '__main__'    :
    main()
