import os
import pandas as pd


from .biomechanics import ranges_of_motion, find_stance_len
from .config import movement_names
from .generate_excel import generate_excel
from .generate_pdf import generate_pdf


def load_data(file: str) -> dict:
    pd_data = pd.read_csv(file, header=None)
    data = {}

    n_swings = find_stance_len(pd_data, "left"), find_stance_len(pd_data, "right")
    for level in ("kinematics", "moment", "power"):
        for side in ("left", "right"):
            for plane in ("sagittal", "frontal", "transversal"):
                names = movement_names[level][plane]
                for index in range(0 if side == "left" else 1, len(names), 2):
                    name = names[index]["en"]
                    data[name] = ranges_of_motion(pd_data, name, side, n_swings)

    return data


def process_files(files: list[str, ...]):
    for file in files:
        save_folder = os.path.dirname(file)
        save_file_no_extension = os.path.splitext(os.path.basename(file))[0]

        data = load_data(file)

        generate_pdf(data, f"{save_folder}/{save_file_no_extension}.pdf")
        generate_excel(data, f"{save_folder}/{save_file_no_extension}.xlsx")
