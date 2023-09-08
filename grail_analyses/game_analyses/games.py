# -*- coding: utf-8 -*-
"""
Created on Tue Aug  1 14:26:25 2023

@author: Florence
"""
from abc import ABC, abstractmethod


class GameAbstract(ABC):
    name: str  # Used for the name of the game
    save_name: str  # Used for the name of the file

    def __init__(self, name: str, save_name: str):
        self.name = name
        self.save_name = save_name

    def __str__(self):
        return self.name

    def __repr__(self):
        return f"*{self.save_name}*"

    @abstractmethod
    def results(self, data, date: str):
        """
        This method should call first the "_common_results" and add their own result
        to the output
        """

    def _common_results(self, date):
        return "Jeu", self.name, "Date", date


def get_games_list() -> tuple[GameAbstract, ...]:
    return (
        AWalkAcrossTheBoard(),
        ItalianAlps(),
        Microbes(),
        PerturbationTrainer(),
        RopeBridge(),
    )


class AWalkAcrossTheBoard(GameAbstract):
    def __init__(self):
        super(AWalkAcrossTheBoard, self).__init__(name="A Walk Across The Board", save_name="WalkAcrossTheBoard")

    def results(self, data, date: str):
        duration = data[-1, 0] * 60
        distance = data[-1, 7]
        max_speed = max(data[:, 6])
        mean_speed = (data[:, 6]).mean()

        return (
            *self._common_results(date),
            ("Duree_s", duration),
            ("Distance_m", distance),
            ("VitesseMax_m/s", max_speed),
            ("VitesseMoy_m/s", mean_speed)
        )


class ItalianAlps(GameAbstract):
    def __init__(self):
        super(ItalianAlps, self).__init__(name="Italian alps", save_name="ItalianAlps")

    def results(self, data, date):
        duration = data[-1, 0] * 60
        distance = data[-1, 2]

        max_speed = max(data[:, 1])
        mean_speed = (data[:, 1]).mean()
        max_inclination = max(data[:, 5])
        min_inclination = min(data[:, 5])
        mean_inclination = (data[:, 5]).mean()

        return (
            *self._common_results(date),
            ("Duree_s", duration),
            ("Distance_m", distance),
            ("VitesseMax_m/s", max_speed),
            ("VitesseMoy_m/s", mean_speed),
            ("InclinaisonMax_deg", max_inclination),
            ("InclinaisonMin_deg", min_inclination),
            ("InclinaisonMoy_deg", mean_inclination)
        )


class Microbes(GameAbstract):
    def __init__(self):
        super(Microbes, self).__init__(name="Microbes", save_name="Microbes")

    def results(self, data, date):
        duration = data[-1, 0] * 60
        distance = data[-1, 2]
    
        max_speed = max(data[:, 1])
        mean_speed = (data[:, 1]).mean()
    
        return (
            *self._common_results(date),
            ("Duree_s", duration),
            ("Distance_m", distance),
            ("VitesseMax_m/s", max_speed),
            ("VitesseMoy_m/s", mean_speed),
        )
    

class PerturbationTrainer(GameAbstract):
    def __init__(self):
        super(PerturbationTrainer, self).__init__(name="Perturbation trainer", save_name="PerturbationTrainer")

    def results(self, data, date):
        duration = data[-1, 0] * 60
        distance = data[-1, 2]
    
        max_speed = max(data[:, 6])
        mean_speed = (data[:, 6]).mean()
    
        return (
            *self._common_results(date),
            ("Duree_s", duration),
            ("Distance_m", distance),
            ("VitesseMax_m/s", max_speed),
            ("VitesseMoy_m/s", mean_speed),
        )


class RopeBridge(GameAbstract):
    def __init__(self):
        super(RopeBridge, self).__init__(name="Rope bridge", save_name="RopeBridge")

    def results(self, data, date):
        duration = data[-1, 0] * 60
        distance = data[-1, 2]

        max_speed = max(data[:, 1])
        mean_speed = (data[:, 1]).mean()
        max_inclination = max(data[:, 5])
        min_inclination = min(data[:, 5])
        mean_inclination = (data[:, 5]).mean()

        return (
            *self._common_results(date),
            ("Duree_s", duration),
            ("Distance_m", distance),
            ("VitesseMax_m/s", max_speed),
            ("VitesseMoy_m/s", mean_speed),
            ("InclinaisonMax_deg", max_inclination),
            ("InclinaisonMin_deg", min_inclination),
            ("InclinaisonMoy_deg", mean_inclination)
        )
