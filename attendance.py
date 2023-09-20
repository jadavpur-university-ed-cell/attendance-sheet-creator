import random

import pandas as pd

names_df = pd.read_excel("names_list.xlsx")
names_list = names_df["Name"].tolist()
names_list = list(set(names_list))

session_titles = input(
    "Enter the titles of the session you want to generate attendance for (comma separated): "
).split(",")
min_attendance = int(input("Enter the minimum attendance count needed: "))
max_attendance = int(input("Enter the maximum attendance count allowed: "))

# session_titles = ["Grant Writing"]
all_depts = [
    "Chemical Engineering",
    "Power Engineering",
    "Mechanical Engineering",
    "Production Engineering",
    "Electrical Engineering",
    "ETCE",
    "Computer Science",
    "Metallurgy Engineering",
    "Construction Engineering",
    "IT",
    "IEE",
    "Food Technology",
    "Printing Engineering",
]
all_years = ["UG-3", "UG-4"]

for title in session_titles:
    num_students = random.randint(min_attendance, max_attendance)
    random_students_list = random.sample(names_list, num_students)
    if title == "Chem Engineering" or title == "Power Engineering":
        dept_list = [title for i in range(num_students)]
        yr_list = ["UG-1" for i in range(num_students)]
    else:
        dept_list = [random.choice(all_depts) for i in range(num_students)]
        yr_list = [random.choice(all_years) for i in range(num_students)]
    sr_num_list = [i for i in range(1, num_students + 1)]
    data = {
        "Sr No.": sr_num_list,
        "Name of Student": random_students_list,
        "Year of Study": yr_list,
        "Department": dept_list,
    }
    df = pd.DataFrame(data)
    df.to_excel(f"./Sheets/{title} Attendance.xlsx", index=False)
