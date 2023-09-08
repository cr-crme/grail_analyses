import os

from .config import grail_save_folder
from .generate_excel import generate_excel


def check_patient_id(patient_id):
    # Check that the patient_id is actually valid
    for root, subfolders, files in os.walk(grail_save_folder):
        if patient_id not in subfolders:
            raise RuntimeError(f"{patient_id} is not valid")


def export_results(patient_id):
    check_patient_id(patient_id)
    save_file = f"{grail_save_folder}/{patient_id}/Record Data/Analyse/{patient_id}_ResumeIntervention.xlsx"
    generate_excel(patient_id, save_file)
