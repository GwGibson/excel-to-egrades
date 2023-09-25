import argparse
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd
import pyperclip


def get_arguments():
    parser = argparse.ArgumentParser(
        description="Enter a scaling factor and optional file path."
    )
    parser.add_argument(
        "--scale",
        default=1,
        type=float,
        help="The scaling factor for the grades.",
    )
    parser.add_argument(
        "--file",
        default=None,
        type=str,
        help="The optional file path for the Moodle excel document.",
    )
    args = parser.parse_args()
    return args.scale, args.file


def acquire_file(optional_file_path=None):
    if optional_file_path:
        return optional_file_path
    Tk().withdraw()
    script_dir = os.path.dirname(os.path.realpath(__file__))
    file_path = askopenfilename(
        initialdir=script_dir,
        title="Please select the moodle excel document containing the grades.",
    )
    return file_path


def parse_data_frame(file_path):
    data_frame = pd.read_excel(file_path)
    id_column = data_frame.columns[2]
    grade_column = data_frame.columns[7]
    data_frame = data_frame[[id_column, grade_column]]
    data_frame[id_column] = data_frame[id_column].str.upper()
    data_frame[id_column] = "txt" + data_frame[id_column]
    return data_frame


def parse_file_paths(input_file_path):
    input_file_name = os.path.splitext(os.path.basename(input_file_path))[0]
    output_file_path = os.path.join(
        os.path.dirname(input_file_path), f"inject_grades_{input_file_name}.js"
    )
    return output_file_path


def create_js_code(data_frame, output_file_path, scaling_factor=1):
    js_code = ""
    for _, row in data_frame.iterrows():
        user_id = row[data_frame.columns[0]]
        raw_grade = row[data_frame.columns[1]]
        grade = 0 if raw_grade == "-" else float(raw_grade) / scaling_factor
        grade = (
            format(grade, ".2f").rstrip("0").rstrip(".")
            if isinstance(grade, float)
            else grade
        )
        js_code += f'document.getElementById("{user_id}").value = "{grade}";\n'

    with open(output_file_path, "w", encoding="utf-8") as file:
        file.write(js_code)
    pyperclip.copy(js_code)


def print_success_message(output_file_path):
    print(
        "Conversion of grades to JavaScript code was successful and the "
        "code has been copied to your clipboard.\n"
        f"The generated JavaScript file can be found at: {output_file_path}"
    )


if __name__ == "__main__":
    factor, path = get_arguments()
    path = acquire_file(path)
    df = parse_data_frame(path)
    output_path = parse_file_paths(path)
    create_js_code(df, output_path, factor)
    print_success_message(output_path)
