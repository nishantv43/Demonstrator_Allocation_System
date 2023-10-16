import pyomo.environ as pyo
from pyomo.environ import *
from pyomo.opt import SolverFactory
import pandas as pd
#from random import random
import time


''' Read the excel sheets into the python variables for
student type and job data'''
dataUG = pd.read_excel(r"D:\Projects\Final Demonstrator Python Code\Demonstrator_Data.xlsx"
                        ,sheet_name="UG")
dataPG = pd.read_excel(r"D:\Projects\Final Demonstrator Python Code\Demonstrator_Data.xlsx"
                        ,sheet_name="PG")
dataR = pd.read_excel(r"D:\Projects\Final Demonstrator Python Code\Demonstrator_Data.xlsx"
                        ,sheet_name="R")
dataJobs = pd.read_excel(r"D:\Projects\Final Demonstrator Python Code\Demonstrator_Data.xlsx"
                        ,sheet_name="Job")
dataCosts = pd.read_excel(r"D:\Projects\Final Demonstrator Python Code\Demonstrator_Data.xlsx"
                        ,sheet_name="Cost")


U = [u for u in dataUG.id]
P = [p for p in dataPG.id]
R = [r for r in dataR.id]
J = [j for j in dataJobs.JobId]

''' Cost of allocating each type of student (UG, PG and R)
as read from the excel sheet of dataCosts'''
cost_u = dataCosts.Cost[0]
cost_p = dataCosts.Cost[1]
cost_r = dataCosts.Cost[2]

''' Define skills data '''
student_skills_UG = {}
student_skills_PG = {}
student_skills_R = {}
job_skills = {}

'''Take the skills from Excel sheet and split by delimeter ','
For eg: Excel has skills like ML,BA
'''
for _, row in dataUG.iterrows():
    student_id = row['id']
    skills = [skill.strip() for skill in row['Skills'].split(',')]
    student_skills_UG[student_id] = skills

for _, row in dataPG.iterrows():
    student_id = row['id']
    skills = [skill.strip() for skill in row['Skills'].split(',')]
    student_skills_PG[student_id] = skills

for _, row in dataR.iterrows():
    student_id = row['id']
    skills = [skill.strip() for skill in row['Skills'].split(',')]
    student_skills_R[student_id] = skills

for _, row in dataJobs.iterrows():
    job_id = row['JobId']
    skills = [skill.strip() for skill in row['Skills_Reqd'].split(',')]
    job_skills[job_id] = skills

''' Define the vacancy of each job '''
ug_vacancy = dataJobs.set_index('JobId')['UG_Vacancy'].to_dict()
pg_vacancy = dataJobs.set_index('JobId')['PG_Vacancy '].to_dict()
r_vacancy = dataJobs.set_index('JobId')['R_Vacancy'].to_dict()

''' Define the days of the week availability '''
UG_demonstrator_availability = {}
PG_demonstrator_availability = {}
R_demonstrator_availability = {}
job_availability  = {}


'''Take the days of the week avalaibilyt from Excel sheet and split by delimeter ','
For eg: Excel has days_availability like Monday, Friday
'''
for _, row in dataUG.iterrows():
    student_id = row['id']
    availability = set(row['Availability'].split(', '))
    UG_demonstrator_availability[student_id] = availability

for _, row in dataPG.iterrows():
    student_id = row['id']
    availability = set(row['Availability'].split(', '))
    PG_demonstrator_availability[student_id] = availability

for _, row in dataR.iterrows():
    student_id = row['id']
    availability = set(row['Availability'].split(', '))
    R_demonstrator_availability[student_id] = availability

for _, row in dataJobs.iterrows():
    job_id = row['JobId']
    availability = set(row['Job_availability '].split(', '))
    job_availability[job_id] = availability

model = pyo.ConcreteModel()

''' Create the three binary decision variables for
allocating UG, PG and R'''
model.x = pyo.Var(U, J,   within=Binary)
x = model.x

model.y = pyo.Var(P, J,   within=Binary)
y = model.y

model.z = pyo.Var(R, J,   within=Binary)
z = model.z


'''Define the objective function - Max Assigment'''
model.objective = Objective(
    expr=sum( x[i, j] for i in U for j in J)
    + sum( y[i, j] for i in P for j in J)
    + sum( z[i, j] for i in R for j in J),
    sense = maximize
)

'''If random allocation is needed please comment the above objective function code
and uncomment below objective function with random().
Also uncomment the import random library on the top of the file'''
# model.objective = Objective(
#     expr=sum( x[d, j] * random() for d in U for j in J)
#     + sum( y[d, j] * random() for d in P for j in J)
#     + sum( z[d, j] * random() for d in R for j in J),
#     sense = maximize
# )

'''DEFINE THE CONSTRAINTS'''

''' 1) Budget Constraint '''
model.budget_constraint = ConstraintList()
for j in J:
    model.budget_constraint.add(
        sum(cost_u * x[d, j] for d in U)
        + sum(cost_p * y[d, j] for d in P)
        + sum(cost_r * z[d, j] for d in R)
        <= dataJobs.Budget[j-1]
    )

''' 2) Required No of Demonstrator for each job Constraint
 For eg: a lecturer may specify that they require two
 postgraduate students '''

model.assignment_constraint = ConstraintList()
for j in J:
        model.assignment_constraint.add(
            sum(x[u, j] for u in U) <= ug_vacancy[j]
        )
        model.assignment_constraint.add(
            sum(y[d, j] for d in P) <= pg_vacancy[j]
        )
        model.assignment_constraint.add(
            sum(z[d, j] for d in R) <= r_vacancy[j]
        )

''' 3.1) Jobs cant be worked  by U constraints '''
model.job_constraint = ConstraintList()
for d in U:
    model.job_constraint.add(
        x[d, 2] == 0
    )
    model.job_constraint.add(
        x[d, 3] == 0
    )
    model.job_constraint.add(
        x[d, 4] == 0
    )

''' 3.2) Jobs cant be worked by P constraints '''
model.job_constraint2 = ConstraintList()
for d in P:
    model.job_constraint2.add(
        y[d, 4] == 0
    )

''' 4) One student can be assigned only one job '''
model.one_job_constraintU = ConstraintList()
for d in U:
    model.one_job_constraintU.add(
            sum(x[d, j] for j in J)  <= 1
        )

model.one_job_constraintP = ConstraintList()
for d in P:
    model.one_job_constraintP.add(
            sum(y[d, j] for j in J)  <= 1
        )

model.one_job_constraintR = ConstraintList()
for d in R:
    model.one_job_constraintR.add(
            sum(z[d, j] for j in J)  <= 1
        )

''' 5) Skills Constraint - if a student is assigned then
he/she should have atleast one common skill
from the required skills for that job '''

model.skills_constraint = ConstraintList()
for d in U:
    for j in J:
        has_common_skill = sum(x[d, j] for s in student_skills_UG.get(d, [])
                               if s in job_skills.get(j, []))
        model.skills_constraint.add(has_common_skill >= x[d, j])

for d in P:
    for j in J:
        has_common_skill = sum(y[d, j] for s in student_skills_PG.get(d, [])
                               if s in job_skills.get(j, []))
        model.skills_constraint.add(has_common_skill >= y[d, j])

for d in R:
    for j in J:
        has_common_skill = sum(z[d, j] for s in student_skills_R.get(d, [])
                               if s in job_skills.get(j, []))
        model.skills_constraint.add(has_common_skill >= z[d, j])

''' 6) Hour Constraint - Hours/day '''
model.hour_constraint_u = ConstraintList()
for d in U:
    model.hour_constraint_u.add(
        sum(dataJobs.Hours[j - 1]  * x[d, j] for j in J)
        <= dataUG.Max_Hours[d - 1]
    )

model.hour_constraint_p = ConstraintList()
for d in P:
    model.hour_constraint_p.add(
        sum(dataJobs.Hours[j - 1]  * y[d, j] for j in J)
        <= dataPG.Max_Hours[d - 1]
    )

model.hour_constraint_r = ConstraintList()
for d in R:
    model.hour_constraint_r.add(
        sum(dataJobs.Hours[j - 1]  * z[d, j] for j in J)
        <= dataR.Max_Hours[d - 1]
    )

''' 7) Availibilty on days of the week for both job and demonstrator'''
model.day_constraint_U = ConstraintList()

for j in J:
    job_days = job_availability[j]

    for d in U:
        demonstrator_days = UG_demonstrator_availability[d]

        if job_days.issubset(demonstrator_days) or demonstrator_days.issuperset(job_days):
            model.day_constraint_U.add(x[d, j] <= 1)
        else:
            model.day_constraint_U.add(x[d, j] == 0)

model.day_constraint_P = ConstraintList()
for j in J:
    job_days = job_availability[j]

    for d in P:
        demonstrator_days = PG_demonstrator_availability[d]

        if job_days.issubset(demonstrator_days) or demonstrator_days.issuperset(job_days):
            model.day_constraint_P.add(y[d, j] <= 1)
        else:
            model.day_constraint_P.add(y[d, j] == 0)

model.day_constraint_R = ConstraintList()
for j in J:
    job_days = job_availability[j]

    for d in R:
        demonstrator_days = R_demonstrator_availability[d]

        if job_days.issubset(demonstrator_days) or demonstrator_days.issuperset(job_days):
            model.day_constraint_R.add(z[d, j] <= 1)
        else:
            model.day_constraint_R.add(z[d, j] == 0)

###############################

start_time = time.time()

solver = SolverFactory('glpk')
results = solver.solve(model)#, tee = True)


end_time = time.time()

time_taken = end_time - start_time
print(f"Time taken to find the optimal solution: {time_taken} seconds")

''' Print the results if optimal solution is found  '''
if results.solver.termination_condition == TerminationCondition.optimal:
    for d in U:
        for j in J:
            if x[d, j].value == 1:
                print(f"Assign undergraduate student D{d} to job {j}")
    print("\n")
    for d in P:
        for j in J:
            if y[d, j].value == 1:
                print(f"Assign postgraduate student P{d} to job {j}")
    print("\n")
    for d in R:
        for j in J:
            if z[d, j].value == 1:
                print(f"Assign research fellow R{d} to job {j}")

    num_allocations_ug = sum(pyo.value(model.x[d, j]) for d in U for j in J)
    num_allocations_pg = sum(pyo.value(model.y[d, j]) for d in P for j in J)
    num_allocations_r = sum(pyo.value(model.z[d, j]) for d in R for j in J)
    total_allocated = num_allocations_ug + num_allocations_pg + num_allocations_r

    print("\nNumber of UG students allocated :", num_allocations_ug)
    print("Number of PG students allocated :", num_allocations_pg)
    print("Number of R students allocated :", num_allocations_r)
    print("\nTotal Number of Allocations :", total_allocated)

    ''' Print the budget value '''
    budget_value = sum(
        cost_u * x[d, j].value
        for d in U for j in J
    ) + sum(
        cost_p * y[d, j].value
        for d in P for j in J
    ) + sum(
        cost_r * z[d, j].value
        for d in R for j in J
    )
    total = sum(num for num in dataJobs.Budget if isinstance(num, int))
    print("\nTotal Available Budget : ", total)
    print("Total Allocated Budget :", budget_value)
else:
    print("No optimal solution found.")
