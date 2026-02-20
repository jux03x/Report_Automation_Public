#Get Input for new report
import yaml
from datetime import datetime
import utils.common_decorater as dc

#Load config.yaml
with open("config.yaml", "r") as f:
    config = yaml.safe_load(f)

#Validate the date format
def get_valid_date(prompt):
    while True:
        date_str = input(prompt)
        try:
            #Check if the date input is valid
            date_obj = datetime.strptime(date_str, "%d.%m.%Y")
            return date_obj  # Return date if correct
        except ValueError:
            print("Invalid format! Please use given format: dd.mm.yyyy.")
            print("Your Input Type is:", type(prompt))

#Validate the datatype float input
def get_valid_float(prompt):
    while True:
        user_input = input(prompt)
        try:
            return float(user_input)
        except ValueError:
            print("Invalid input! Please enter a valid decimal number (use '.').")
            print("Your Input Type is:", type(prompt))

#Validate the datatype int input
def get_valid_int(prompt):
    while True:
        user_input = input(prompt)
        try:
            return int(user_input)
        except ValueError:
            print("Invalid input! Please enter a valid integer.")
            print("Your Input Type is:", type(prompt))

#Validate the datatype int or default input for downtime
def get_valid_int_default_downtime(prompt):
    while True:
        user_input = input(prompt)
        try:
            return int(user_input)
        except ValueError:
            print("Default selected")
            return 0
        
#Validate the datatype float or default input for system
def get_valid_float_default_system(prompt):
    while True:
        user_input = input(prompt)
        try:
            return float(user_input)
        except ValueError:
            print("Default selected")
            return config["system"]["availability"]


#Function for input
dc.logger_advc
def ask_input():
    print("=== Capture new report data ===")
    print("##############################################################################################\n")
    print("Please enter the requested Format\n")

    date_start = get_valid_date("Start of Report period (dd.mm.yyyy): ")
    date_end = get_valid_date("End of Report period (dd.mm.yyyy): ")


    print("##############################################################################################\n")
    print("Please use '.' for decimal numbers\n")

    storage = get_valid_float(("Storage: "))
    tenant_count = get_valid_int(("Tenant-count: "))
    document_count = get_valid_int(("Document-count: "))

    sys1_downtime = get_valid_int_default_downtime(("\nService Downtime PROD (Minuten) - Hit enter for Default(0): ") or 0)
    sys2_downtime = get_valid_int_default_downtime(("Service Downtime TEST (Minuten) - Hit enter for Default(0): ") or 0)

    print("\n________________________________________________________________________________________\n")

    #Get the value from config.yaml
    avg_accesstime = config["system"]["availability"]
    print(str(avg_accesstime) + " is Default for Monthly avarage access time in sec")
    avg_accesstime = get_valid_float_default_system(("Monthly avarage access time in sec - Hit enter for Default: ") or avg_accesstime)
    print("\n#########################################################################################\n")

    file=str(input("Filename in Inputpath: "))
    print("\n#########################################################################################\n")

    #Write Data to a Temporary input.yaml for later use
    input_data = {
        "report_period": {
            "year": int(date_start.year),
            "date_start": date_start, 
            "date_end": date_end},
        "values": {
            "storage": storage,
            "document_count": document_count,
            "tenant_count": tenant_count
        },
        "downtime": {
            "sys1_minutes": sys1_downtime,
            "sys2_minutes": sys2_downtime
        },
        "system": {
            "availablitiy": avg_accesstime
        },
        "filename": file
    }

    # Save for late, in case of a retry
    with open("input.yaml", "w") as f:
        yaml.dump(input_data, f)

    return input_data
