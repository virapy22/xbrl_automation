# xbrl_automation
This project was developed to verify the accuracy of reported the financial information presented on the mandatory XBRL regulatory returns.

XBRL Mapping & Comparison
Developed by virapy22

## Prerequisites

### System

- Anaconda: For Windows Python 3.7 version. After installation is complete, please search for Anaconda prompt in the Start and pin this to you taskbar for ease of access. 

### Files
- Ensure all required files are in the same directory as the code files
- Do not rename the files or tabs within the sheets, the program is designed to read data using those names only
- In the mapping, files new rows may be added for new line items however please do not add any new columns, the program is only designed to read from the specified file location(s).
- If you'd like to replace the template with a new blank one, please copy `master_template.xlsx` from the blank_template_for_replacement folder into the xbrl_program folder and rename it to `template.xlsx`. Please delete the prior template before doing this. 

### Data
- Changes to new cells or existing cell mapping may be made only by appending new rows. The column wise data may NOT be changed as adding new columns exceeds the buffer limit. 
- Each cell location reference in the mapping must start with the applicable sign (+/-), should no operator apply, please use the '+' operator by default. 
- Operations within cell location reference: the program can carry out basic operations between multiple cells from various locations. 
to manipulate cells within the same reference sheet separate all location references by column (ie. A1+B7-C6 becomes +A1,+B7,-C6)
to manipulate cells in different reference sheets, enter the cell reference under the applicable sheet names.
- When typing out cell location references please ensure there are no spaces at point. Spaces are parsed as blanks, hence rendering the data non-existent. 


## Execution

This code is divided into 2 parts: Step1.py and Step2.py. 

To run these you will have to follow the following steps:

- Open anaconda prompt through the start menu. 
- Ensure all excel files pertaining to this program (mapping, fs, xbrl_form, xbrl_form_mapping, template) are closed before proceeding further.
- Ensure the directory is set to the location of the code and excel files. you can change that by entering the following command followed by enter.

`cd <filepath>`

- Once the default path is set to the file path, run the following command. This runs the step1.py script which will then load all the data from the map and then parse it step by step and then replace the data. 

`python step1.py`

Note: this program is designed to give you the run log in the prompt window. 

An 'operation successful' message will appear should the program run successfully. This means all the FS data has been appended into the template successfully. 

- The second part of the program runs the step2.py script which loads all the data from the xbrl map and compares the template to the xbrl form line by line. Running the following command will execute this:

`python step2.py`

Note: this program is designed to give you the run log in the prompt window. 

A 'Comparison complete' message will appear should the program run successfully. This means all the data has been compared and all outliers have been highlighted in red in both the template and xbrl_form excel files.

- You may enter `exit()` into the prompt window to exit safely. 

### Error Handling and Support requests
Please contact me for any support requests as this code was designed in line with particular financial information and may not scale perfectly to other FS.
