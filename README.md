# Demonstrator Allocation System

This project aims to solve the problem of allocating demonstrators to various jobs in an educational institution using Integer Linear Programming (ILP). The system utilizes the Pyomo library to formulate and solve the optimization problem and can be customized for different scenarios.

## Installation

Before running the Demonstrator Allocation System, you need to install the required dependencies as per Windows:

### Python

Python is the programming language used for this project. Make sure you have Python 3 installed.

# Check if Python is installed
python3 --version

# If python not installed, Follow below instructions for Windows:

Visit the official Python website: https://www.python.org/downloads/
Download the latest version of Python for Windows.
Run the installer executable.
Make sure to check the "Add Python to PATH" option during installation.
Follow the installation prompts.

# Install Pandas library
pip install pandas

# Install pyomo
pip install pyomo

## Solvers
This project to work with smaller dataset use glpk as the solver. However, if you want to run with large dataset you have use the gurobi installed.

# GLPK
GLPK is an open-source linear programming solver. Install it using the following steps:

1. open the web page - https://sourceforge.net/projects/winglpk/
2. Click on download for glpk
2. Unzip it and paste in the C: drive
4. Now set the environment variables
    - Click on System Settings
    - Inside that click on advanced system settings
    - A pop up window open click on the tab "Advanced"
    - Now click on "Environment variables"
    - Another pop up window opens
    - Now double click on "Path" in the bottom section of "System   Variables"
    - Now set the new path as "C:\glpk-4.65\w64" and save


# Gurobi
Gurobi is a commercial optimization solver that can be used as an alternative to GLPK for larger problems. To install Gurobi, you need to obtain a license from Gurobi website (This project uses this solver with Academic Free 1 year license). After obtaining a license, you can install it using pip:

You can obtain Gurobi from the official website: https://www.gurobi.com/

You can find the installation guide in the website: https://www.gurobi.com/documentation/10.0/quickstart_windows/py_using_pip_to_install_gr.html

Install Gurobi by following the provided installation instructions.

# Replace 'YOUR_LICENSE_KEY' with your actual Gurobi license key
pip install gurobipy
grbgetkey YOUR_LICENSE_KEY


## Excel Data

Python program read the data from excel "Demostrator_data" attached in the zipped folder. The path in all three models to locate the excel should be updated as per your machine path. Otherwise, you will get file path not found error.

Excel containts following 5 spreadsheets :
    UG - undergradute data with id should be unique value
    PG - postgraduate data with id should be unique value
    R - research data with id should be unique value
    Job - job details with with id as unique value
    Cost - cost of assigning the type of demonstrator.
    JobsPerClass - Jobs associated with classes

# Important Points

You can edit the student details or add new student in UG, PG and R spreadsheets. But the id should be unique.

For basic and extension model to run :

1. Dont remove the jobs. Jobs has to be 4.
2. You can edit the skills, Hours, Job_availability, vacancy, budget to meaningful values. Follow the format as defined already the sheets for the corresponding columns.
3. However, you can add more jobs with id being the unique value

For Advanced Model to run:

1. Dont remove the jobs. Jobs has to be 12 with 3 classes.
2. You can edit the skills, Hours, Job_availability, vacancy, budget to meaningful values. Follow the format as defined already the sheets for the corresponding columns.
4. However, you can add more jobs with more classes with ids being the unique value.

