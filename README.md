# Automated Time Table Scheduling for IIIT Dharwad

This project implements an **Automated Timetable Scheduling System** for IIIT Dharwad.  
It generates academic and exam timetables automatically using input datasets (faculty, classrooms, courses, students, constraints) and scheduling logic.

---

## 📁 Folder Structure

- `data/` — Contains input CSV/JSON files:  
  - `classroom_data.csv`  
  - `course_data.csv`  
  - `exam_data.csv`  
  - `invigilator_data.csv`  
  - `student_data.csv`  
  - `faculty_availability.csv`  
  - `constraints.json`  

- `timetable_automation/` — Core source code for timetable generation  
- `tests/` — Unit tests for modules  
- `docs/` — Project documentation (DPR, design, reports, diagrams)  
- `.gitignore` — Specifies files/folders ignored by Git  
- `requirements.txt` — (if present) Lists dependencies  

---

## Run
```
python -m timetable_automation.main
```

## Test
```
python -m unittest discover tests
```

## Team Name
BumbleBee


## Team Members
- Raskin Verma
- Vishwa D
- Sindhu Talari
- Udit Dadhich
