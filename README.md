# Automated Time Table Scheduling for IIIT Dharwad

A Python-based automated timetable generation system for universities that handles multiple departments, elective courses, faculty constraints, and room allocations with intelligent scheduling algorithms.

## Team Information

**Team Name:** BumbleBee

**Team Members:**
- Raskin Verma
- Vishwa D
- Sindhu Talari
- Udit Dadhich

## Table of Contents

- [Features](#features)
- [Project Structure](#project-structure)
- [Installation](#installation)
- [Input Data Format](#input-data-format)
- [Configuration](#configuration)
- [Usage](#usage)
- [Output Files](#output-files)
- [Algorithm Details](#algorithm-details)
- [Testing](#testing)
- [Documentation](#documentation)
- [Contributing](#contributing)

## Features

- **Multi-Department Support**: Generates timetables for multiple departments (CSE, DSAI, ECE, etc.) simultaneously
- **Semester-wise Scheduling**: Supports first-half and second-half semester courses
- **Elective Course Management**: Intelligently schedules elective courses with room assignments
- **Faculty Constraint Handling**: Prevents faculty scheduling conflicts across departments
- **Smart Room Allocation**: 
  - Separate handling for classrooms and labs
  - Prevents room conflicts across all departments
  - Persistent room assignment for courses
- **Session Type Support**: Handles Lectures (L), Tutorials (T), and Practicals (P)
- **Break Management**: Automatically excludes break slots (morning, lunch, evening)
- **Visual Excel Output**: 
  - Color-coded timetables with merged cells
  - Automatic column width adjustment
  - Detailed course legend with faculty and room information
- **Faculty Timetables**: Generates individual timetables for each faculty member across all departments
- **Deterministic Scheduling**: Uses random seed for reproducible results

## Project Structure

```
Automated-Time-Table-Scheduling-for-IIIT-Dharwad/
├── data/
│   ├── CSE_courses.csv          # CSE department courses
│   ├── DSAI_courses.csv         # DSAI department courses
│   ├── ECE_courses.csv          # ECE department courses
│   ├── Faculty.csv              # Faculty information (optional)
│   ├── rooms.csv                # Available rooms and labs
│   └── timeslots.csv            # Time slot definitions
├── docs/
│   └── BumbleBee_DPR.pdf        # Detailed Project Report
├── tests/
│   ├── test_cases.md            # Test case documentation
│   ├── test_course.py           # Course class tests
│   ├── test_elective_allotment.py
│   ├── test_end_to_end.py
│   ├── test_scheduler_basic.py
│   └── test_session_allocation.py
├── timetable_automation/
│   ├── __init__.py
│   └── main.py                  # Core scheduling logic
├── main.py                      # Entry point script
└── README.md
```

## Installation

### Prerequisites

- Python 3.7 or higher
- pip package manager


## Input Data Format

### 1. Course Files (`data/*_courses.csv`)

Each department should have its own CSV file with the following columns:

| Column | Description | Example |
|--------|-------------|---------|
| `Course_Code` | Unique course identifier | `CS301`, `DSAI201` |
| `Course_Title` | Full course name | `Data Structures`, `Machine Learning` |
| `Faculty` | Faculty name(s), use `/` for multiple options | `Dr. Smith/Dr. Jones` |
| `L-T-P-S-C` | Lecture-Tutorial-Practical-Self Study-Credits | `3-1-2-0-4` |
| `Semester_Half` | `1` for first half, `2` for second half, `0` for both | `1`, `2`, or `0` |
| `Elective` | `1` if elective course, `0` if core | `0` or `1` |


### 2. Rooms File (`data/rooms.csv`)

Defines available classrooms and labs:

| Column | Description | Example |
|--------|-------------|---------|
| `Room_ID` | Unique room identifier | `CR101`, `Lab1` |
| `Type` | `classroom` or `lab` | `classroom`, `lab` |


### 3. Time Slots File (`data/timeslots.csv`)

Defines all time slots for the timetable:

| Column | Description | Example |
|--------|-------------|---------|
| `Start_Time` | Slot start time | `09:00` |
| `End_Time` | Slot end time | `10:30` |


## Configuration

### Excluded Time Slots

The system automatically excludes break periods defined in `main.py`:

```python
self.excluded_slots = ["07:30-09:00", "10:30-10:45", "13:15-14:00", "17:30-18:30"]
```

These slots are reserved for:
- Early morning (before official start)
- Morning break
- Lunch break
- Evening break

### Adding New Departments

To add a new department, update the `departments` dictionary in `main.py`:

```python
departments = {
    "CSE": "data/CSE_courses.csv",
    "DSAI": "data/DSAI_courses.csv",
    "ECE": "data/ECE_courses.csv",  # Add new department
}
```
### Random Seed

For reproducible results, the system uses a fixed random seed:

```python
RANDOM_SEED = 42
```

Change this value to generate different timetable variations.

## Usage

### Basic Usage

Run the main script from the project root directory:

```bash
python main.py
```

This will generate timetables for all configured departments.

### Running Individual Tests

```bash
# Run all tests
python -m pytest tests/

# Run specific test file
python -m pytest tests/test_scheduler_basic.py

# Run with verbose output
python -m pytest tests/ -v
```

## Output Files

The system generates the following Excel files:

### 1. Student Timetables

**Files:** `CSE_timetable.xlsx`, `DSAI_timetable.xlsx`, etc.

Each file contains:
- **First_Half Sheet**: Timetable for first-half semester courses
- **Second_Half Sheet**: Timetable for second-half semester courses

**Features:**
- Color-coded courses for easy identification
- Merged cells for multi-slot sessions
- Course legend with:
  - Course code and title
  - Faculty name
  - Assigned classroom/lab
  - Color indicator
- Elective courses listed separately with room assignments

### 2. Faculty Timetables

**File:** `faculty_timetable.xlsx`

Contains individual sheets for each faculty member showing:
- All teaching assignments across departments
- Course details with room numbers
- Color-coded for visual clarity
- Merged cells for continuous sessions

## Algorithm Details

### Scheduling Strategy

1. **Course Prioritization**:
   - Non-elective courses scheduled first
   - Elective courses handled with special logic
   - One representative elective selected per sheet

2. **Session Allocation**:
   - Lectures (L): 1.5-hour sessions
   - Tutorials (T): 1-hour sessions
   - Practicals (P): 2-hour sessions (max one per day)

3. **Constraint Handling**:
   - Faculty cannot teach multiple sessions simultaneously
   - Rooms cannot be double-booked
   - Labs restricted to one practical per day
   - Automatic gap insertion between merged sessions

4. **Room Assignment**:
   - Persistent room assignment per course
   - Global room conflict prevention across departments
   - Separate pools for classrooms and labs
   - Elective courses assigned conflict-free rooms

5. **Retry Mechanism**:
   - Maximum 10 attempts per session allocation
   - Randomized day selection for better coverage

### Elective Course Handling

The system uses a unique approach for electives:
1. Schedules one representative elective in the timetable
2. Tracks all elective options per sheet
3. Computes conflict-free room assignments for all electives
4. Displays all options in the legend with assigned rooms

## Testing

The project includes comprehensive test coverage:

- **`test_course.py`**: Tests Course class initialization and parsing
- **`test_scheduler_basic.py`**: Basic scheduler functionality
- **`test_session_allocation.py`**: Session allocation logic
- **`test_elective_allotment.py`**: Elective course scheduling
- **`test_end_to_end.py`**: Complete workflow testing

Run tests to verify system functionality before deploying.

## Documentation

For detailed project information, refer to:
- **Detailed Project Report**: `docs/BumbleBee_DPR.pdf`
- **Test Cases**: `tests/test_cases.md`

## Key Features Explained

### Global Room Conflict Prevention

The system maintains a `global_room_usage` dictionary that tracks room occupancy across all departments, preventing double-booking scenarios.

### Faculty Workload Management

Multiple faculty options can be specified for a course using the `/` separator. The system randomly selects one for scheduling, allowing flexibility in faculty assignment.

### Color-Coded Output

Each course is automatically assigned a unique color from a predefined palette, making timetables easy to read and visually appealing.

### Automatic Column Width Adjustment

Excel output automatically adjusts column widths based on content length for optimal readability.

## Troubleshooting

### Common Issues

**Issue**: Courses not being scheduled

**Solution**: 
- Check `excluded_slots` don't overlap with available time
- Ensure sufficient rooms in `rooms.csv`
- Verify faculty availability constraints

**Issue**: Room conflicts appearing

**Solution**:
- Ensure `global_room_usage` is properly shared across schedulers
- Check that room types match course requirements (classroom vs lab)

**Issue**: Faculty overloaded

**Solution**:
- Add alternative faculty using `/` separator in course file
- Reduce number of courses per faculty
- Adjust `MAX_ATTEMPTS` for better distribution


## Acknowledgments

Special thanks to IIIT Dharwad for the project opportunity and guidance.

---

**For questions or support**, please open an issue on the GitHub repository or contact any of the team members.