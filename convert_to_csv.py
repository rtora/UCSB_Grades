"""
Converts grade.xlsx into CSV files under /csv

Directory structure is:
csv/
    - [year]_[quarter]/
        - [major]_[course_number]/
            - [instructor].csv

"""
import os
import re
import csv
import shutil
from collections import defaultdict

from xlsx2csv import Xlsx2csv


def main():
    print("Cleaning csv/...")
    shutil.rmtree("csv/", ignore_errors=True)

    print("Converting grades.xlsx to tmp csv files...")
    converter = Xlsx2csv("grades.xlsx")
    converter.convert("tmp/", sheetid=0)

    print("Parsing tmp csv files...")
    for file in os.listdir("tmp/"):
        quarter, year = re.search(r"([A-z]+)\s(\d+)\.csv", file).groups()
        quarter_directory = f"csv/{year}_{quarter.upper()}/"
        os.makedirs(quarter_directory, exist_ok=True)
        print(f"Outputting to {quarter_directory}...")
        with open(f"tmp/{file}") as quarter_data:
            reader = csv.reader(quarter_data)
            convert_quarter(year, quarter, quarter_directory, reader)

    print("Cleaning tmp/...")
    shutil.rmtree("tmp/")


def convert_quarter(year, quarter, quarter_directory, csv_data):
    data = defaultdict(lambda: defaultdict(list))
    for level, course, instructor, grade, count in list(csv_data)[1:]:
        instructor = instructor.strip().replace(" ", "_")
        if len(instructor) == 0:
            instructor = "UNKNOWN_INSTRUCTOR"
        major, course_number = re.search(r"([A-z]*)\s+(\d*)", course).groups()
        formatted_course = f"{major}_{course_number}"
        data[formatted_course][instructor].append(
            [year, quarter, level, major, course_number, instructor, grade, count]
        )

    for course in data:
        for instructor in data[course]:
            dir_path = f"{quarter_directory}{course}"
            os.makedirs(dir_path, exist_ok=True)
            path = f"{dir_path}/{instructor}.csv"
            with open(path, "w") as f:
                writer = csv.writer(f)
                writer.writerows(data[course][instructor])


if __name__ == "__main__":
    main()
