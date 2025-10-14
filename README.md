# Automated Timetable Generator

A comprehensive Python-based system for generating academic timetables for universities/colleges with multiple departments, semesters, and faculty members.

---

## Table of Contents

- [Features](#features)  
- [Requirements](#requirements)  
  - [Python Dependencies](#python-dependencies)  
  - [Input Files](#input-files)  
- [Configuration](#configuration)  
  - [Time Settings](#time-settings)  
  - [Break Times](#break-times)  
- [Usage](#usage)  
  - [Basic Usage](#basic-usage)  
  - [Functions](#functions)  
- [Output Files](#output-files)  
- [Key Features Explained](#key-features-explained)  
  - [Scheduling Constraints](#scheduling-constraints)  
  - [Priority Scheduling](#priority-scheduling)  
- [Troubleshooting](#troubleshooting)  
- [Advanced Configuration](#advanced-configuration)  
- [Notes](#notes)  
- [Support](#support)  
- [License](#license)  
- [Team Name & Members](#team-name--members)

---

## Features

- **Multi-Department Support**: Generate timetables for multiple departments and semesters simultaneously.  
- **Intelligent Scheduling**: Automatically schedules lectures, tutorials, labs, and self-study sessions.  
- **Faculty Management**: Tracks faculty availability and prevents conflicts.  
- **Room Allocation**: Smart room assignment based on capacity and type (lecture halls, computer labs, hardware labs).  
- **Basket Courses**: Supports elective/basket courses with flexible scheduling.  
- **Break Management**: Automatically schedules breaks including staggered lunch breaks.  
- **Excel Output**: Generates color-coded, professional Excel timetables.  
- **Conflict Detection**: Identifies and reports unscheduled courses with detailed reasons.  
- **Faculty Timetables**: Creates individual timetables for each faculty member.

---

## Requirements

### Python Dependencies

```bash
pip install pandas openpyxl
```

> Add additional dependencies as required by your implementation (e.g., `xlsxwriter`, `numpy`) if used.

### Input Files

Place the following CSV files in the same directory as the script:

#### `combined.csv` (Course information)

Columns:

- `Department` — Department code (e.g., `CS`, `EC`)  
- `Semester` — Semester number (e.g., `3`, `4`, `5`)  
- `Course Code` — Unique course identifier (e.g., `CS301`, `B1-CS151`)  
- `Course Name` — Full course name  
- `Faculty` — Faculty name(s) — use `/` for multiple options  
- `L` — Lecture credits (used to calculate session count)  
- `T` — Tutorial hours per week  
- `P` — Practical/Lab hours per week  
- `S` — Self-study hours per week  
- `C` — Total course credits  
- `total_students` — Number of students enrolled  
- `Schedule` — `Yes`/`No` (optional, defaults to `Yes`)

Example:
```csv
Department,Semester,Course Code,Course Name,Faculty,L,T,P,S,C,total_students,Schedule
CS,3,CS301,Data Structures,Dr. Smith,3,1,2,0,4,65,Yes
EC,4,EC401,Digital Electronics,Dr. Jones,2,0,4,0,4,45,Yes
CS,5,B1-CS501,Machine Learning,Dr. Brown,3,0,0,2,3,35,Yes
```

#### `rooms.csv` (Room information)

Columns:

- `id` — Unique room identifier  
- `roomNumber` — Room number/name  
- `capacity` — Maximum student capacity  
- `type` — Room type (e.g., `LECTURE_ROOM`, `COMPUTER_LAB`, `HARDWARE_LAB`, `SEATER_120`, `SEATER_240`)

Example:
```csv
id,roomNumber,capacity,type
R001,A101,60,LECTURE_ROOM
R002,A102,70,LECTURE_ROOM
R003,Lab1,35,COMPUTER_LAB
R004,Lab2,35,COMPUTER_LAB
R005,Hall1,120,SEATER_120
R006,Hall2,240,SEATER_240
```

---

## Configuration

### Time Settings

Edit constants at the top of the script to suit your institute’s timings:

```python
START_TIME = time(9, 0)          # Day start time
END_TIME = time(18, 30)          # Day end time
LECTURE_DURATION = 3             # 1.5 hours (in 30-min slots)
LAB_DURATION = 4                 # 2 hours (in 30-min slots)
TUTORIAL_DURATION = 2            # 1 hour (in 30-min slots)
SELF_STUDY_DURATION = 2          # 1 hour (in 30-min slots)
```

> Note: The implementation assumes a 30-minute base timeslot.

### Break Times

```python
LUNCH_WINDOW_START = time(12, 30)  # Lunch can start from
LUNCH_WINDOW_END = time(14, 0)     # Lunch must end by
LUNCH_DURATION = 60                # Lunch duration in minutes
```

---

## Usage

### Basic Usage

Run the main script (example file name `timetable_generator.py`):

```bash
python timetable_generator.py
```

This will generate:

1. `timetable_all_departments.xlsx` — Combined timetable for all departments  
2. `all_faculty_timetables.xlsx` — Individual schedules for all faculty  
3. `unscheduled_courses.xlsx` — Report of any courses that couldn't be scheduled

### Functions

#### `generate_all_timetables()`
Creates a single Excel file with separate sheets for each department-semester combination.

#### `check_unscheduled_courses()`
Analyzes the generated timetable and reports courses that weren't fully scheduled according to their L-T-P-S requirements.

#### `generate_faculty_timetables()`
Creates comprehensive timetables for each faculty member showing all their teaching assignments.

---

## Output Files

### `timetable_all_departments.xlsx`
- **Overview Sheet**: Index of department-semester combinations  
- **Department Sheets**: One sheet per department-semester containing:
  - Color-coded timetable grid (Days × Time slots)
  - Course details (code, type, room, faculty)
  - Legend showing course colors and details
  - Self-study only courses list
  - Unscheduled components (if any)

### `all_faculty_timetables.xlsx`
- **Overview Sheet**: List of all faculty with total class count  
- **Faculty Sheets**: Individual schedules for each faculty showing:
  - Day-wise schedule
  - Course details
  - Room assignments
  - Department-semester information

### `unscheduled_courses.xlsx`
Report containing:
- Courses that couldn't be fully scheduled  
- Required vs. scheduled hours (L-T-P-S)  
- Missing components  
- Possible reasons for scheduling failure

---

## Key Features Explained

### Scheduling Constraints

1. **Faculty Constraints**
   - Maximum 2 course components per day (configurable).  
   - Minimum 3-hour gap between classes (except lab after lecture/tutorial).  
   - No scheduling conflicts across assigned slots.

2. **Room Allocation**
   - Automatic selection based on student count.  
   - Adjacent lab rooms for large lab batches when needed.  
   - Priority room booking rules differentiate lectures vs labs.

3. **Break Management**
   - Morning break example: 10:30–10:45 (15 minutes).  
   - Staggered lunch breaks for different semesters.  
   - Inter-class breaks maintained to avoid back-to-back overload.

4. **Basket Courses**
   - Parallel scheduling of elective courses.  
   - Shared time slots across basket groups.  
   - Individual room allocation per course.

### Priority Scheduling

Courses are scheduled in the following priority order:

1. Regular course labs (highest priority)  
2. Lectures with high credit hours  
3. Tutorials  
4. Basket course components (lowest priority)

---

## Troubleshooting

### Common Issues & Solutions

**Issue**: "No suitable room found"  
- **Solution**: Increase room count in `rooms.csv` or adjust room capacities.

**Issue**: "Faculty overbooked"  
- **Solution**: Add more faculty options using `/` separator in `combined.csv` or reduce course load.

**Issue**: "Lab sessions not scheduled"  
- **Solution**: Ensure sufficient lab rooms of correct `type` (e.g., `COMPUTER_LAB`, `HARDWARE_LAB`).

**Issue**: "Courses marked as unscheduled"  
- **Solution**: Check `unscheduled_courses.xlsx` for detailed reasons and adjust:
  - Add more rooms
  - Distribute faculty load
  - Adjust semester course count

---

## Advanced Configuration

### Custom Color Palette

Modify the `COLOR_PALETTE` list in the script to change course colors:

```python
COLOR_PALETTE = [
    "FF5733", "33FF57", "3357FF", ...
]
```

### Room Type Customization

Add new room types in `rooms.csv` and update the `find_suitable_room()` function to support them.

### Section Management

Student sections are automatically calculated from the `total_students` column:

- Default maximum: 70 students per section (configurable)  
- Large classes are split into multiple sections automatically

---

## Notes

- **Slot Duration**: All time slots are 30 minutes.  
- **Working Days**: Monday to Friday (5 days).  
- **Course Credits**: Lecture credits determine number of lecture sessions.  
- **Self-Study Only**: Courses with only `S` hours (no `L/T/P`) are listed separately.  
- **Basket Course Format**: Use formats such as `B1-COURSECODE`, `B2-COURSECODE`, etc.

---

## Support

If you encounter issues:

1. Check `unscheduled_courses.xlsx` for scheduling problems.  
2. Verify that input CSV formats match the requirements above.  
3. Review console output for warnings and errors.  
4. Optionally, open an issue on the repository with the CSV samples and error messages.

---

## License

Provided as-is for educational and administrative use. (Add a proper license file if needed, e.g., `LICENSE` with MIT/Apache/GPL text.)

---

## Team Name & Members

**Team Name:** BumbleBee

**Team Members:**
- Raskin Verma  
- Vishwa D  
- Sindhu Talari  
- Udit Dadhich
