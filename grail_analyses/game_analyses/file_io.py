import os

from .config import grail_save_folder
from .generate_excel import generate_excel


def check_patient_id(patient_id):
    # Check that the patient_id is actually valid
    for root, subfolders, files in os.walk(grail_save_folder):
        if patient_id in subfolders:
            return
    raise RuntimeError(f"{patient_id} is not valid")


def export_results(patient_id):
    check_patient_id(patient_id)
    generate_excel(patient_id, f"{grail_save_folder}/{patient_id}/Record Data")
