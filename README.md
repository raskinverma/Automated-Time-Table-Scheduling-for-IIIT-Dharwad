## Automated Timetable Generation System
# Overview

This project automates the generation of academic timetables for multiple departments (e.g., CSE, DSAI, ECE, etc.) using input data on courses, faculty, rooms, and time slots.

The system reads structured .csv files and produces:

Department-wise student timetables (Excel format)

A combined faculty timetable workbook (Excel format)

The generated files include color formatting, merged cells for continuous sessions, and a legend summarizing course–faculty–room mappings.

## Key Features

Automatic lecture, tutorial, and lab scheduling
Conflict-free faculty and room allocation
Elective session management across departments
Common global room usage to prevent overlaps
Color-coded Excel outputs with legends
Combined faculty-wise timetable generation

## Project Structure

AUTOMATED-TIME-TABLE/
├── data/
│ ├── CSE_courses.csv
│ ├── DSAI_courses.csv
│ ├── ECE_courses.csv
│ ├── Faculty.csv
│ ├── rooms.csv
│ ├── timeslots.csv
│
├── timetable_automation/
│ └── main.py      # Main script (contains Scheduler class)
│
├── tests/
│ ├── test_data/     # Test CSVs & output samples
│ ├── test_course.py
│ ├── test_end_to_end.py
│ ├── test_scheduler_basic.py
│ └── test_session_allocation.py
│
├── docs/
│ ├── BumbleBee_DPR.pdf # Project report
│ └── main.tex      # LaTeX source
│
└── README.md

## Prerequisites

Install Python dependencies:

pip install pandas openpyxl

## Input Files

All data files are stored in the data/ folder.

File	Description
*_courses.csv	Course information per department
Faculty.csv	Faculty details
rooms.csv	Room ID, type (lab/classroom), and capacity
timeslots.csv	Start and end times for each slot
Example: CSE_courses.csv
Course_Code	Course_Title	Faculty	L-T-P-S-C	Semester_Half	Elective
CS101	Programming in C	John Doe	3-1-0-0-4	1	0
CS202	DBMS Lab	Jane Doe	0-0-3-0-2	2	0
CSEL01	Elective Course	Smith / Alan	3-0-0-0-3	1	1
Example: timeslots.csv
Start_Time	End_Time
09:00	10:00
10:00	11:00
11:00	12:00
...	...
Example: rooms.csv
Room_ID	Type	Capacity
CR-101	Classroom	60
CR-102	Classroom	60
LAB-201	Lab	30

## Code Structure

The scheduling logic is implemented in timetable_automation/main.py and built around two main classes:

Course

Handles parsing of individual course data (code, title, faculty, and L–T–P structure).

Scheduler

Manages:

Room & faculty allocation

Slot assignment

Excel sheet formatting and color coding

Generating both student and faculty timetable workbooks

Each department is handled independently, but a shared global_room_usage map ensures no inter-department room conflicts.

## Execution

To run the timetable generator:

python timetable_automation/main.py


## The script will:

Read department course CSVs (e.g., CSE_courses.csv, DSAI_courses.csv)

Load room and slot data

Generate two timetables per department:

CSE_timetable.xlsx

DSAI_timetable.xlsx

Merge and produce one combined faculty timetable:

faculty_timetable.xlsx

## Output
 Student Timetable (<DEPT>_timetable.xlsx)

Each department gets a workbook with two sheets:

First_Half (Semester half 1 or 0)

Second_Half (Semester half 2 or 0)

Each sheet includes:

Timetable grid with courses, rooms, and sessions

Color-coded cells for each course

A legend mapping course codes → faculty → room

 Faculty Timetable (faculty_timetable.xlsx)

A combined workbook with one sheet per faculty, showing:

All assigned sessions

Continuous slots merged

Unique colors for each course

## Testing

Tests are included in the /tests directory.

Run all tests using:

pytest tests/

## Notes

Ensure all CSVs are well-formatted (no missing L–T–P–S–C values).

Elective courses must have Elective = 1 in the CSV.

The script uses a fixed RANDOM_SEED = 42 for reproducibility.

Excluded time slots (like breaks) are defined within the code under:

self.excluded_slots = ["07:30-09:00", "10:30-10:45", "13:15-14:00", "17:30-18:30"]

## Team BumbleBee
Raskin Verma	
Vishwa D
Sindhu Talari	
Udit Dadhich