import math


def ranges_of_motion(file_data, variable, side, length):
    if side not in ("left", "right"):
        raise ValueError("Wrong side")

    data = file_data[file_data.iloc[:, 1] == variable]
    data_side = data[data.iloc[:, 0] == side]
    average = data_side.iloc[:, 3:].mean()

    swing_phase = average[length[0 if side == "left" else 1]:]
    maximum_swing = max(swing_phase)
    minimum_swing = min(swing_phase)
    range_swing = maximum_swing - minimum_swing

    stance_phase = average[:length[0 if side == "left" else 1]]
    maximum_stance = max(stance_phase)
    minimum_stance = min(stance_phase)
    range_stance = maximum_stance - minimum_stance

    return {
        "Maximum Oscillation": maximum_swing,
        "Minimum Oscillation": minimum_swing,
        "Range Oscillation": range_swing,
        "Maximum Appui": maximum_stance,
        "Minimum Appui": minimum_stance,
        "Range Appui": range_stance,
    }


def find_stance_len(data, side) -> int:
    if side not in ("left", "right"):
        raise ValueError("Wrong side")

        # Find number of columns for swing and stance phase
    swing_time = data[data.iloc[:, 1] == f"{'L' if side == 'left' else 'R'}.Swing.Time"]
    stance_time = data[data.iloc[:, 1] == f"{'L' if side == 'left' else 'R'}.Stance.Time"]

    sided_swing_time = swing_time[swing_time.iloc[:, 0] == side]
    sided_stance_time = stance_time[stance_time.iloc[:, 0] == side]

    swing_time_average = sided_swing_time.iloc[:, 3].mean()
    stance_time_average = sided_stance_time.iloc[:, 3].mean()

    total_time_average = swing_time_average + stance_time_average
    swing_percent = swing_time_average / total_time_average * 100
    return math.ceil(swing_percent * 101 / 100)
