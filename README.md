### CAN_Matrix data conversion into YAML

#### Description:
There are two approaches for parsing data from `CAN Matrix.xlsx` file.
Ang covering these into YAML file `output.yaml` file.

Both approaches either `routine_appr.py` or `compact_appr.py` lead to 
the same `output.yaml` file. 

#### Prepare & Launch Python scripts:
To clone the data from `CAN Matrix.xlsx` and convert the data to `output.yml` file it's required to:
- clone the repository: `https://github.com/smart2004/CAN_Matrix.git`;
- may install virtual environment: `python -m venv venv` and run it: `source venv/scripts/activate`[for Windows] or `source venv/bin/activate`[for Linux];
- run the command in bash terminal(to fulfill the required libraries): `pip install -r requirements.txt`;
- and run the scripts as: `python routine_appr.py` /OR/ `python compact_appr.py`.


Thanks & Regards,
smart200481