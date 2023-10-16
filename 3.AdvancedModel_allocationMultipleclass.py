import pyomo.environ as pyo
from pyomo.environ import *
from pyomo.opt import SolverFactory
import pandas as pd
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
dataClassesJobs = pd.read_excel(r"D:\Projects\Final Demonstrator Python Code\Demonstrator_Data.xlsx"
                        ,sheet_name="JobsPerClass")

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

for _, row in dataClassesJobs.iterrows():
    job_id = row['JobId']
    skills = [skill.strip() for skill in row['Skills_Reqd'].split(',')]
    job_skills[job_id] = skills

C = list(set(dataClassesJobs.Class))
JP = [jp for jp in dataClassesJobs.JobId]


''' Group jobs by class'''
class_jobs = dataClassesJobs.groupby('Class')

''' Create sets of jobs for each class'''
class_job_sets = {
    c: class_jobs.get_group(c)['JobId'].tolist()
    for c in C
}

class_budget = {}
for c in C:
    class_budget[c] = sum(dataClassesJobs[dataClassesJobs['Class'] == c]['Budget'])


class_vacancy = {}
for c in C:
    class_vacancy[c] = sum(dataClassesJobs[dataClassesJobs['Class'] == c]['Vacancy'])

class_skills_reqd = {}
for c in C:
    skills = dataClassesJobs[dataClassesJobs['Class'] == c]['Skills_Reqd']
    skills_list = [skill.strip() for skills_row in skills for skill in skills_row.split(',')]
    class_skills_reqd[c] = skills_list

class_hours = {}
for c in C:
    class_hours[c] = sum(dataClassesJobs[dataClassesJobs['Class'] == c]['Hours'])

''' Define the vacancy of each job '''
ug_vacancy = dataClassesJobs.set_index('JobId')['UG_Vacancy'].to_dict()
pg_vacancy = dataClassesJobs.set_index('JobId')['PG_Vacancy '].to_dict()
r_vacancy = dataClassesJobs.set_index('JobId')['R_Vacancy'].to_dict()

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

for _, row in dataClassesJobs.iterrows():
    job_id = row['JobId']
    availability = set(row['Job_availability '].split(', '))
    job_availability[job_id] = availability

model = pyo.ConcreteModel()

''' Create the three binary decision variables for
allocating UG, PG and R'''
model.x = pyo.Var(U, C, JP, within=Binary)
x = model.x

model.y = pyo.Var(P, C, JP, within=Binary)
y = model.y

model.z = pyo.Var(R, C, JP, within=Binary)
z = model.z


''' Assign the allocation to 0 for those jobs which doesnt belong in the class'''
for u in U:
    for c in C:
        for j in JP:
            if j not in class_job_sets[c]:
                x[u, c, j].value = 0

for u in P:
    for c in C:
        for j in JP:
            if j not in class_job_sets[c]:
                y[u, c, j].value = 0

for u in R:
    for c in C:
        for j in JP:
            if j not in class_job_sets[c]:
                z[u, c, j].value = 0

'''Define the objective function - Max Assigment'''
model.objective = pyo.Objective(
    expr=sum(x[u, c, j] for u in U for c in C for j in class_job_sets[c])
    + sum(y[u, c, j] for u in P for c in C for j in class_job_sets[c])
    + sum(z[u, c, j] for u in R for c in C for j in class_job_sets[c]),
    sense=pyo.maximize
)

'''DEFINE THE CONSTRAINTS'''

''' 1) Budget Constraint to consider the class budget'''
model.budget_constraint = pyo.ConstraintList()
for c in C:
    for j in class_job_sets[c]:
        model.budget_constraint.add(
            sum(x[u, c, j] * cost_u  for u in U)
            + sum(cost_p * y[d, c, j] for d in P)
            + sum(cost_r * z[d,c, j] for d in R)
            <= dataClassesJobs.Budget[j-1]
        )

''' 2) Required No of Demonstrator for each job Constraint
 For eg: a lecturer may specify that they require two
 postgraduate students '''
model.assignment_constraint = ConstraintList()
for c in C:
    for j in class_job_sets[c]:
        model.assignment_constraint.add(
            sum(x[u, c, j] for u in U) <= ug_vacancy[j]
        )
        model.assignment_constraint.add(
            sum(y[d, c, j] for d in P) <= pg_vacancy[j]
        )
        model.assignment_constraint.add(
            sum(z[d, c, j] for d in R) <= r_vacancy[j]
        )

ug_job_restriction = [2, 3, 4, 6, 7, 8, 10, 11, 12]
''' 3.1) Jobs cant be worked  by U constraints '''
model.job_constraint = ConstraintList()
for c in C:
    for d in U:
        for j in ug_job_restriction:
                model.job_constraint.add(x[d, c, j] == 0)

pg_job_restriction = [4, 8, 12]
''' 3.2) Jobs cant be worked by P constraints '''
model.job_constraint2 = ConstraintList()
for c in C:
    for d in P:
        for j in pg_job_restriction:
            model.job_constraint2.add(y[d, c, j] == 0)

''' 4) One student can be assigned only one job '''
model.one_job_constraintU = ConstraintList()
for u in U:
    model.one_job_constraintU.add(
        sum(x[u, c, j] for c in C for j in class_job_sets[c]) <= 1
    )

model.one_job_constraintP = ConstraintList()
for u in P:
    model.one_job_constraintP.add(
        sum(y[u, c, j] for c in C for j in class_job_sets[c]) <= 1
    )

model.one_job_constraintR = ConstraintList()
for u in R:
    model.one_job_constraintR.add(
        sum(z[u, c, j] for c in C for j in class_job_sets[c]) <= 1
    )

''' 5) Skills Constraint - if a student is assigned then
he/she should have atleast one common skill
from the required skills for that job '''

model.skills_constraint = ConstraintList()
for d in U:
    for c in C:
        for j in class_job_sets[c]:
            has_common_skill = sum(
                x[d, c, j] for s in student_skills_UG.get(d, [])
                if s in job_skills.get(j, [])
            )
            model.skills_constraint.add(has_common_skill >= x[d, c, j])

for d in P:
    for c in C:
        for j in class_job_sets[c]:
            has_common_skill = sum(
                y[d, c, j] for s in student_skills_PG.get(d, [])
                if s in job_skills.get(j, [])
            )
            model.skills_constraint.add(has_common_skill >= y[d, c, j])

for d in R:
    for c in C:
        for j in class_job_sets[c]:
            has_common_skill = sum(
                z[d, c, j] for s in student_skills_R.get(d, [])
                if s in job_skills.get(j, [])
            )
            model.skills_constraint.add(has_common_skill >= z[d, c, j])

''' 6) Hour Constraint - Hours/day '''
model.hour_constraint_u = pyo.ConstraintList()
for u in U:
        model.hour_constraint_u.add(
            sum(x[u, c, j] * dataClassesJobs.Hours[j-1] for c in C
                for j in class_job_sets[c] )
            <= dataUG.Max_Hours[u - 1]
        )

model.hour_constraint_p = pyo.ConstraintList()
for u in P:
        model.hour_constraint_p.add(
            sum(y[u, c, j] * dataClassesJobs.Hours[j-1] for c in C
                for j in class_job_sets[c] )
            <= dataPG.Max_Hours[u - 1]
        )

model.hour_constraint_r = pyo.ConstraintList()
for u in R:
        model.hour_constraint_r.add(
            sum(z[u, c, j] * dataClassesJobs.Hours[j-1] for c in C
                for j in class_job_sets[c] )
            <= dataR.Max_Hours[u - 1]
        )

''' 7) Availibilty on days of the week for both job and demonstrator'''
model.day_constraint_U = ConstraintList()
for c in C:
    for j in class_job_sets[c]:
        job_days = job_availability[j]
        for d in U:
            demonstrator_days = UG_demonstrator_availability[d]
            if job_days.issubset(demonstrator_days) or demonstrator_days.issuperset(job_days):
                model.day_constraint_U.add(x[d,c, j] <= 1)
            else:
                model.day_constraint_U.add(x[d,c, j] == 0)

model.day_constraint_P = ConstraintList()

for c in C:
    for j in class_job_sets[c]:
        job_days = job_availability[j]
        for d in P:
            demonstrator_days = PG_demonstrator_availability[d]
            if job_days.issubset(demonstrator_days) or demonstrator_days.issuperset(job_days):
                model.day_constraint_P.add(y[d,c, j] <= 1)
            else:
                model.day_constraint_P.add(y[d,c, j] == 0)

model.day_constraint_R = ConstraintList()
for c in C:
    for j in class_job_sets[c]:
        job_days = job_availability[j]
        for d in R:
            demonstrator_days = R_demonstrator_availability[d]
            if job_days.issubset(demonstrator_days) or demonstrator_days.issuperset(job_days):
                model.day_constraint_R.add(z[d, c, j] <= 1)
            else:
                model.day_constraint_R.add(z[d, c, j] == 0)

start_time = time.time()

# Solve the model
solver = SolverFactory('glpk')
results = solver.solve(model)#,  tee=True)

# Print the results
print("----- Results -----")

end_time = time.time()

time_taken = end_time - start_time
print(f"Time taken to find the optimal solution: {time_taken} seconds")

''' Print the results if optimal solution is found  '''
if results.solver.termination_condition == TerminationCondition.optimal:
    print("Undergraduate student assignments:")
    for d in U:
        for c in C:
            for j in JP:
                if pyo.value(x[d, c, j]) == 1:
                    print(f"Student U{d} assigned to Class {c}, Job {j}")

    print("\nPostgraduate student assignments:")
    for d in P:
        for c in C:
            for j in JP:
                if pyo.value(y[d, c, j]) == 1:
                    print(f"Student P{d} assigned to Class {c}, Job {j}")

    print("\nResearch student assignments:")
    for d in R:
        for c in C:
            for j in JP:
                if pyo.value(z[d, c, j]) == 1:
                    print(f"Student R{d} assigned to Class {c}, Job {j}")

    print("Objective value:", pyo.value(model.objective))
    print("Available Budget per Class:")
    for c in C:
        print(f"Class {c}: {class_budget[c]}")

    ''' Print the budget value '''
    print("\nAllocated Budget per Class:")
    for c in C:
        allocated_budget = sum(
            cost_u * x[u, c, j].value
            for u in U for j in class_job_sets[c]
        ) + sum(
            cost_p * y[p, c, j].value
            for p in P for j in class_job_sets[c]
        ) + sum(
            cost_r * z[r, c, j].value
            for r in R for j in class_job_sets[c]
        )
        print(f"Class {c}: {allocated_budget}")
else:
    print("No optimal solution found.")
