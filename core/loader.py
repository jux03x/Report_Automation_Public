import shutil
import yaml
from openpyxl import load_workbook
from utils.dates import format_month_label

#Load config.yaml
with open("config.yaml", "r") as f:
    config = yaml.safe_load(f)
#Load input.yaml
with open("input.yaml", "r") as f:
    input = yaml.safe_load(f)

file=input["filename"]
date_end = input["report_period"]["date_end"]
template_path=config["workbook"]["template_path"]
output_path=config["workbook"]["output_path"]
label = format_month_label(date_end)


def load_and_copy():
    filename = file[:-10] + label + ".xlsx"
    target = output_path + filename
    origin_path = template_path + file

    shutil.copy(origin_path, target)
    wb = load_workbook(target)

    return wb, target
