#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import re
from word import wd

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
                self.bom.write(">|"+self.typeCon+"|"+str(self.nbePin)+"|"+self.tabLigne[0]+"|RESSOURCE|\n")

                #Ecriture des lignes du connecteur#            
                for self.n in self.lst_tmp:

                    #Ecriture de la ligne.
                    self.bom.write(self.n+"\n")
            else:
                print("ERREUR LIB -> "+self.typeCon+"|"+self.con)
                break
                self.bom.close()
                os.remove(self.pathBom)

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
        

def main():        
    bm = library("C:/Python34/Scripts/Doc2Lst/travail/INTERFACE_DE_TEST_CUC05_1A.docx")
    bm.pathBom = "C:/Python34/Scripts/Doc2Lst/out"
    bm.toBom("/INTERFACE_DE_TEST_CUC05_2A")
    bm.quit()

if __name__=="__main__":
    main()
