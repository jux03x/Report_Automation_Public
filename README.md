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
│   ├── dashboard.py        # Sheet processor: dashboard / service report
│   └── system_availability.py  # Sheet processor: service availability
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

7. Follow the script flow. Either type `skip` if `input.yaml` is already filled correctly (jump to step 9), or press `Enter` to start the input manually.

8. Enter the input values as requested by the script.

9. Let the program handle the rest.

10. Check the output folder for mistakes.

11. Done.

12. Repeat from the start every reporting cycle (update the input before each new run).


# Input
.......


# Output
.......

# Timesavings
........
