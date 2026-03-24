# CS1-Swiss-Post
Instructions on file import:

File formatting:

1. Data must be entered into an Excel file that follows the structure of "input_template.xlsx".
2. For the Trafo and PV data:
    2.1 Timestamps must be stored under column "Zeit".
    2.2 Timestamps must be stored in 10 minute intervals.
    2.3 Column title must include "-avg[W]" for power values.
3. Amount of trafo input sheets is arbitrary.
4. For the EV charging profile (LKW):
    4.1 Data must contain a time column in HH:MM:SS format.
    4.2 Data must contain a column with the exact name "Total kW".
    4.3 time column must be in 15-minute resolution covering one full weekday example.
    4.5 Saturday and Sunday are assumed to be 0 load.
5. For the delivery data (Zustellung):
    5.1 Data must be provided as a full-day template (00:00–23:45).
    5.2 First column must contain time values in HH:MM:SS format.
    5.3 One column must exist for each weekday from monday to saturday containing power in kW.
    5.5 Sunday is not included and is assumed to be 0.
    5.6 Data represents winter baseline conditions (summer adjustment is applied in processing).


Configuration:

1. Configurable variables are in a separate worksheet in the input excel file under "config" (as in "input_template.xlsx"). 
2. Only ever change entries in the "values" column. 
3. Expected value types are mentioned in "comments" column in the respective row.


Data Selection:

1. Upon running "main.py", a file dialogue window will be
opened. Select the Excel file with relevant data. The "input_template.xlsx" may be selected to see an example run.
2. A window will open after a short time prompting the user
to select sheets including transformer data. Select all
relevant sheets and click "Apply & Exit".
3. A subsequent window will open prompting the user to select
PV data if available. Select all relevant sheets and click 
"Apply & Exit".
4. The same pattern goes for the PV charging sheet as well as the delivery profile sheet.

