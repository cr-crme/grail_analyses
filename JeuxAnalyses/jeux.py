# -*- coding: utf-8 -*-
"""
Created on Tue Aug  1 14:26:25 2023

@author: Florence
"""

class AWalkAcrossTheBoard():
    def __init__(self):
        self.nom_txt = '*WalkAcrossTheBoard*'  # NOM DANS FICHIER .TXT
        self.nom = "A Walk Across The Board"  # NOM DANS FICHIER .XSLX
    
    def calculs(self, donnees, AAAA, MM, DD):
        duration = donnees[-1, 0] * 60
        distance = donnees[-1, 7]
    
        maxSpeed = max(donnees[:, 6])
        meanSpeed = (donnees[:, 6]).mean()
    
        resultats = (("Jeu", self.nom), 
                     ("Date_AAAA_MM_DD", f"{AAAA}_{MM}_{DD}"), 
                     ("Duree_s", duration), 
                     ("Distance_m", distance),
                     ("VitesseMax_m/s", maxSpeed),
                     ("VitesseMoy_m/s", meanSpeed),
                     )

        return resultats


class ItalianAlps():
    def __init__(self):
        self.nom_txt = '*ItalianAlps*'  # NOM DANS FICHIER .TXT
        self.nom = "Italian alps"  # NOM DANS FICHIER .XSLX
    
    def calculs(self, donnees, AAAA, MM, DD):
        duration = donnees[-1, 0] * 60
        distance = donnees[-1, 2]

        maxSpeed = max(donnees[:, 1])
        meanSpeed = (donnees[:, 1]).mean()
        maxInclination = max(donnees[:, 5])
        minInclination = min(donnees[:, 5])
        meanInclination = (donnees[:, 5]).mean()

        resultats = (("Jeu", self.nom), 
                     ("Date_AAAA_MM_DD", f"{AAAA}_{MM}_{DD}"), 
                     ("Duree_s", duration), 
                     ("Distance_m", distance),
                     ("VitesseMax_m/s", maxSpeed),
                     ("VitesseMoy_m/s", meanSpeed),
                     ("InclinaisonMax_deg", maxInclination),
                     ("InclinaisonMin_deg", minInclination),
                     ("InclinaisonMoy_deg", meanInclination)
                    )

        return resultats


class Microbes():
    def __init__(self):
        self.nom_txt = '*Microbes*'  # NOM DANS FICHIER .TXT
        self.nom = "Microbes"  # NOM DANS FICHIER .XSLX
    
    def calculs(self, donnees, AAAA, MM, DD):
        duration = donnees[-1, 0] * 60
        distance = donnees[-1, 2]
    
        maxSpeed = max(donnees[:, 1])
        meanSpeed = (donnees[:, 1]).mean()
    
        resultats = (("Jeu", self.nom), 
                     ("Date_AAAA_MM_DD", f"{AAAA}_{MM}_{DD}"), 
                     ("Duree_s", duration), 
                     ("Distance_m", distance),
                     ("VitesseMax_m/s", maxSpeed),
                     ("VitesseMoy_m/s", meanSpeed),
                     )

        return resultats
    

class PerturbationTrainer():
    def __init__(self):
        self.nom_txt = '*PerturbationTrainer*'  # NOM DANS FICHIER .TXT
        self.nom = "Perturbation trainer"  # NOM DANS FICHIER .XSLX
    
    def calculs(self, donnees, AAAA, MM, DD):
        duration = donnees[-1, 0] * 60
        distance = donnees[-1, 2]
    
        maxSpeed = max(donnees[:, 6])
        meanSpeed = (donnees[:, 6]).mean()
    
        resultats = (("Jeu", self.nom), 
                     ("Date_AAAA_MM_DD", f"{AAAA}_{MM}_{DD}"), 
                     ("Duree_s", duration), 
                     ("Distance_m", distance),
                     ("VitesseMax_m/s", maxSpeed),
                     ("VitesseMoy_m/s", meanSpeed),
                     )

        return resultats


class RopeBridge():
    def __init__(self):
        self.nom_txt = '*RopeBridge*'  # NOM DANS FICHIER .TXT
        self.nom = "Rope bridge"  # NOM DANS FICHIER .XSLX
    
    def calculs(self, donnees, AAAA, MM, DD):
        duration = donnees[-1, 0] * 60
        distance = donnees[-1, 2]

        maxSpeed = max(donnees[:, 1])
        meanSpeed = (donnees[:, 1]).mean()
        maxInclination = max(donnees[:, 5])
        minInclination = min(donnees[:, 5])
        meanInclination = (donnees[:, 5]).mean()

        resultats = (("Jeu", self.nom), 
                     ("Date_AAAA_MM_DD", f"{AAAA}_{MM}_{DD}"), 
                     ("Duree_s", duration), 
                     ("Distance_m", distance),
                     ("VitesseMax_m/s", maxSpeed),
                     ("VitesseMoy_m/s", meanSpeed),
                     ("InclinaisonMax_deg", maxInclination),
                     ("InclinaisonMin_deg", minInclination),
                     ("InclinaisonMoy_deg", meanInclination)
                    )

        return resultats
