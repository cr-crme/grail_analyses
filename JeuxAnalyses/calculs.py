# -*- coding: utf-8 -*-
"""
Created on Tue Aug  1 14:52:09 2023

@author: Florence
"""


def AWalkAcrossTheBoard(nom, data, AAAA, MM, DD):
    duration = data[-1, 0] * 60
    distance = data[-1, 7]

    maxSpeed = max(data[:, 6])
    meanSpeed = (data[:, 6]).mean()

    resultats = [nom,
                 f"{AAAA}_{MM}_{DD}",
                 duration,
                 distance,
                 maxSpeed,
                 meanSpeed,
                 ]  # Resume des variables

    return resultats


def ItalianAlps(nom, data, AAAA, MM, DD):
    duration = data[-1, 0] * 60
    distance = data[-1, 2]

    maxSpeed = max(data[:, 1])
    meanSpeed = (data[:, 1]).mean()
    maxInclination = max(data[:, 5])
    minInclination = min(data[:, 5])
    meanInclination = (data[:, 5]).mean()

    resultats = [nom,
                 f"{AAAA}_{MM}_{DD}",
                 duration,
                 distance,
                 maxSpeed,
                 meanSpeed,
                 maxInclination,
                 minInclination,
                 meanInclination
                 ]  # Resume des variables

    return resultats


# def KiteFlyer(nom, data, AAAA, MM, DD):
#     return resultats


def Microbes(nom, data, AAAA, MM, DD):
    duration = data[-1, 0] * 60
    distance = data[-1, 2]

    maxSpeed = max(data[:, 1])
    meanSpeed = (data[:, 1]).mean()

    resultats = [nom,
                 f"{AAAA}_{MM}_{DD}",
                 duration,
                 distance,
                 maxSpeed,
                 meanSpeed,
                 ]  # Resume des variables

    return resultats


def PerturbationTrainer(nom, data, AAAA, MM, DD):
    duration = data[-1, 0] * 60
    distance = data[-1, 2]

    maxSpeed = max(data[:, 1])
    meanSpeed = (data[:, 1]).mean()

    resultats = [nom,
                 f"{AAAA}_{MM}_{DD}",
                 duration,
                 distance,
                 maxSpeed,
                 meanSpeed,
                 ]  # Resume des variables

    return resultats


def RopeBridge(nom, data, AAAA, MM, DD):
    duration = data[-1, 0] * 60
    distance = data[-1, 2]

    maxSpeed = max(data[:, 1])
    meanSpeed = (data[:, 1]).mean()
    maxInclination = max(data[:, 5])
    minInclination = min(data[:, 5])
    meanInclination = (data[:, 5]).mean()

    resultats = [nom,
                 f"{AAAA}_{MM}_{DD}",
                 duration,
                 distance,
                 maxSpeed,
                 meanSpeed,
                 maxInclination,
                 minInclination,
                 meanInclination
                 ]  # Resume des variables
    return resultats


# def TrafficJam(nom, data, AAAA, MM, DD):
#     return resultats


# def TrafficJamLR(nom, data, AAAA, MM, DD):
#     return resultats
