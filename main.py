import yaml
from pathlib import Path
#from openpyxl import load_workbook

# Main class which executes the whole script
from input import ask_input
#from core.loader import load_and_copy
import utils.common_decorater as cd

# --- load config ---
with open("config.yaml", "r") as f:
    config = yaml.safe_load(f)


@cd.begin_of_func
def main():
    print("\n######################################################################################")
    print("\nSTART of Report Automation\n")
    print("#######################################################################################")
    forward=str(input("\nType 'skip', if you already have the right input in the input.yaml. If not, hit enter to start input: "))
    print("\n")
    
    if forward != "skip":
        #Ask for Input, by calling the function
        input_values = ask_input()
        print("\nInput successfully created: ", input_values)

    # # --- load input ---
    # with open("input.yaml", "r") as f:
    #     input_data = yaml.safe_load(f)


    # print("#########################################################################################")
    # print("Load and Copy of former document")
    # copy_created=load_and_copy()
    # print("Copy was created and is found in the output:", copy_created[1])
    # current_workbook=str(copy_created[1])

    # workbook = load_workbook(current_workbook)

    # # --- save workbook ---
    # workbook.save(current_workbook)
    
    

if __name__ == "__main__":
    main()