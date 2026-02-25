# Report_Automation_Public
This public repository provides a template framework for automating table-based reports. It can be extended and customized to suit a wide range of reporting needs.

Currently, the repository is configured for the example use case shown below. The actual Excel files are intentionally not included in the repository. Each upgrade to an Excel file would result in a new version, and excluding them keeps the repository lightweight and maintainable.

From a cybersecurity perspective, the repository is safe to clone: the code contains only Python scripts, templates, and configuration files. Excel files are excluded because they can contain macros, which may execute arbitrary code and pose security risks. By keeping Excel files out of the public repo, users can safely use and extend the automation scripts without the risk of inadvertently executing malicious macros.

Not that I would ever intentionally do something like this, but this approach makes the repository appear more trustworthy overall. 😊

# Code Structure

```
Report_Automation_Public/
│
├── core/
│   ├── interface.py        # Abstract base interface for sheet processors
│   └── loader.py           # Loads and copies the template workbook
│
├── sheets/
│   ├── base.py             # Base class for single-column sheet logic
│   ├── base2.py            # Base class for double-column sheet logic
│   ├── count_size.py       # Sheet processor: document count & size
│   ├── dashboard.py        # Sheet processor: dashboard 
│   └── system_availability.py  # Sheet processor: availability
│
├── utils/
│   ├── common_decorater.py # Reusable decorators (e.g. logging, timing)
│   └── dates.py            # Date formatting helpers
│
├── input_template/         # Folder for the Excel template file
├── output/                 # Folder for generated report files
│
├── config.yaml             # Sheet configurations (rows, columns, names)
├── input.yaml              # Report input values (dates, counts, downtime)
├── input.py                # CLI input collection script
└── main.py                 # Entry point – orchestrates the full pipeline
```

---

# Getting Started

1. Clone the Repository
2. Use the Repository as it is (Skip to 3.) or customize the Repository for your required use case.
   - Keep the structure and the modularity.
   - Configure the Input (`input.py` & `input.yaml`).
   - Customize the `config.yaml` as needed.
   - Use `utils` or create new ones.
   - Keep `interface.py` & `loader.py`.
   - Customize the `base.py`'s and the sheet scripts as you need them.
   - Call the new sheet scripts in `main.py`.
   > This can take some time depending on the use case. Adapting from a similar project took around 1–2 hours due to many similarities.

3. Start the Report Script.
   - Execute `main.py` with the IDE of your choice.
   - Or execute the script in the command line as shown below.
   > Python must be installed, as well as the required libraries.

4. Open Command Line / PowerShell

5. Change directory to the folder where `main.py` is stored:
   ```bash
   cd C:\path\to\your\Report_Automation_Public
   ```

6. Execute `main.py` with the following command:
   ```bash
   python main.py
   ```

7. Follow the script flow. Either type `skip` if `input.yaml` is already filled correctly (jump to step 10), or press `Enter` to start the input manually.

<img width="1315" height="361" alt="Screenshot 2026-02-25 100302" src="https://github.com/user-attachments/assets/d5cb18dd-6e1e-4cee-9652-08411a434561" />

9. Enter the input values as requested by the script.
<img width="1100" height="886" alt="Screenshot 2026-02-25 100748" src="https://github.com/user-attachments/assets/2370599b-e1f7-4ad8-b877-0ae3f77fcb46" />

11. Let the program handle the rest.

12. Check the output folder for mistakes.

13. Done.

14. Repeat from the start every reporting cycle (update the input before each new run).

!! Keep in mind if you cunduct manuall changes in the file, due to formatting difrences, the next automation may not work anymore. !! 


# Input & Output
Jan >>> Feb

dashboard

<img width="254" height="152,8" alt="image" src="https://github.com/user-attachments/assets/9256f817-2bdd-4536-8163-e87708c745cb" />  --> <img width="251" height="154" alt="image" src="https://github.com/user-attachments/assets/37c783a1-27d2-40c3-83d3-c2eeaf838846" />

system availability

<img width="314,2" height="110" alt="image" src="https://github.com/user-attachments/assets/dfc5c86d-d313-4888-ac9c-6edf9d89647f" /> --> <img width="320" height="110,6" alt="image" src="https://github.com/user-attachments/assets/4e3d97d9-b9c3-481d-b70a-849dd4f635f5" />

count_size

<img width="299,2" height="73,8" alt="image" src="https://github.com/user-attachments/assets/2310657e-6965-4e64-91ab-c4b875c6c958" /> --> <img width="310,2" height="75,4" alt="image" src="https://github.com/user-attachments/assets/694f05b2-0eff-4507-b7da-26c19a469d09" />


Preview Feb >> Mar
Template_Feb26.xlsx is in the input directory now:

<img width="198,2" height="59,6" alt="image" src="https://github.com/user-attachments/assets/d262c5c3-ba54-45e4-8796-92fbd1dbf329" />

dashboard

<img width="251" height="154" alt="image" src="https://github.com/user-attachments/assets/37c783a1-27d2-40c3-83d3-c2eeaf838846" /> --> <img width="238,2" height="152,8" alt="image" src="https://github.com/user-attachments/assets/0a45d7b7-868a-458b-89b5-427b3c454325" />


# Timesavings
The saved time depends on the complexity of the report. In this case you would save around 20 min per report per month. If you have 10 simalar reports you would safe 200 min per month and 40 h in a year.

In one productive usecase this saves me around 2 hours per month. 

Downside is, the more diffrent the reports are the higher is the configuration time.
