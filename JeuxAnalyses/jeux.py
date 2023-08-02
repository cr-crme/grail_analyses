# -*- coding: utf-8 -*-
"""
Created on Tue Aug  1 14:26:25 2023

@author: Florence
"""

class AWalkAcrossTheBoard():
    def init(self):
        nom_txt = '*WalkAcrossTheBoard*'  # NOM DANS FICHIER .TXT
        self.nom = "A Walk Across The Board"  # NOM DANS FICHIER .XSLX
                    
        return nom_txt, self.nom
    
    
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


# class ItalianAlps():
#     def init(self):
#         nom_txt = '*ItalianAlps*'  # NOM DANS FICHIER .TXT
#         self.nom = "Italian alps"  # NOM DANS FICHIER .XSLX
#         variables = ["Jeu",
#                       "Date_AAAA_MM_DD",
#                       "Duree_s",
#                       "Distance_m",
#                       "VitesseMax_m/s",
#                       "VitesseMoy_m/s",
#                       "InclinaisonMax_deg",
#                       "InclinaisonMin_deg",
#                       "InclinaisonMoy_deg",
#                       ]  # NOM VARIABLES
    
#         return nom_txt, self.nom, variables

#     def calculs(self, donnees, AAAA, MM, DD):
#         duration = donnees[-1, 0] * 60
#         distance = donnees[-1, 2]

#         maxSpeed = max(donnees[:, 1])
#         meanSpeed = (donnees[:, 1]).mean()
#         maxInclination = max(donnees[:, 5])
#         minInclination = min(donnees[:, 5])
#         meanInclination = (donnees[:, 5]).mean()

#         resultats = [self.nom,
#                      f"{AAAA}_{MM}_{DD}",
#                      duration,
#                      distance,
#                      maxSpeed,
#                      meanSpeed,
#                      maxInclination,
#                      minInclination,
#                      meanInclination
#                      ]  # Resume des variables

#         return resultats

# # def KiteFlyer():
# #     return nom_txt, nom, variables


# def Microbes():
#     nom_txt = '*Microbes*'  # NOM DANS FICHIER .TXT
#     nom = "Microbes"  # NOM DANS FICHIER .XSLX
#     variables = ["Jeu",
#                  "Date_AAAA_MM_DD",
#                  "Duree_s",
#                  "Distance_m",
#                  "VitesseMax_m/s",
#                  "VitesseMoy_m/s",
#                  ]  # NOM VARIABLES

#     return nom_txt, nom, variables


# def PerturbationTrainer():
#     nom_txt = '*PerturbationTrainer*'  # NOM DANS FICHIER .TXT
#     nom = "Perturbation Trainer"  # NOM DANS FICHIER .XSLX
#     variables = ["Jeu",
#                  "Date_AAAA_MM_DD",
#                  "Duree_s",
#                  "Distance_m",
#                  "VitesseMax_m/s",
#                  "VitesseMoy_m/s",
#                  ]  # NOM VARIABLES

#     return nom_txt, nom, variables


# def RopeBridge():
#     nom_txt = '*RopeBridge*'  # NOM DANS FICHIER .TXT
#     nom = "Rope bridge"  # NOM DANS FICHIER .XSLX
#     variables = ["Jeu",
#                  "Date_AAAA_MM_DD",
#                  "Duree_s",
#                  "Distance_m",
#                  "VitesseMax_m/s",
#                  "VitesseMoy_m/s",
#                  "InclinaisonMax_deg",
#                  "InclinaisonMin_deg",
#                  "InclinaisonMoy_deg",
#                  ]  # NOM VARIABLES

#     return nom_txt, nom, variables


# # def TrafficJam():
# #     return nom_txt, nom, variables


# # def TrafficJamLR():
# #     return nom_txt, nom, variables
