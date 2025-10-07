# Automated Timetable Generator

A comprehensive Python-based system for generating academic timetables for universities/colleges with multiple departments, semesters, and faculty members.

## Features

- **Multi-Department Support**: Generate timetables for multiple departments and semesters simultaneously
- **Intelligent Scheduling**: Automatically schedules lectures, tutorials, labs, and self-study sessions
- **Faculty Management**: Tracks faculty availability and prevents conflicts
- **Room Allocation**: Smart room assignment based on capacity and type (lecture halls, computer labs, hardware labs)
- **Basket Courses**: Supports elective/basket courses with flexible scheduling
- **Break Management**: Automatically schedules breaks including staggered lunch breaks
- **Excel Output**: Generates color-coded, professional Excel timetables
- **Conflict Detection**: Identifies and reports unscheduled courses with detailed reasons
- **Faculty Timetables**: Creates individual timetables for each faculty member

## Requirements

### Python Dependencies
```bash
pip install pandas openpyxl
```

### Input Files

The system requires three CSV files in the same directory:

#### 1. `combined.csv`
Contains course information with the following columns:
- **Department**: Department code (e.g., CS, EC)
- **Semester**: Semester number (e.g., 3, 4, 5)
- **Course Code**: Unique course identifier (e.g., CS301, B1-CS151)
- **Course Name**: Full course name
- **Faculty**: Faculty name(s) - use `/` for multiple options
- **L**: Lecture credits (used to calculate session count)
- **T**: Tutorial hours per week
- **P**: Practical/Lab hours per week
- **S**: Self-study hours per week
- **C**: Total course credits
- **total_students**: Number of students enrolled
- **Schedule**: "Yes"/"No" (optional, defaults to "Yes")

Example:
```csv
Department,Semester,Course Code,Course Name,Faculty,L,T,P,S,C,total_students,Schedule
CS,3,CS301,Data Structures,Dr. Smith,3,1,2,0,4,65,Yes
EC,4,EC401,Digital Electronics,Dr. Jones,2,0,4,0,4,45,Yes
CS,5,B1-CS501,Machine Learning,Dr. Brown,3,0,0,2,3,35,Yes
```

#### 2. `rooms.csv`
Contains room information with the following columns:
- **id**: Unique room identifier
- **roomNumber**: Room number/name
- **capacity**: Maximum student capacity
- **type**: Room type (LECTURE_ROOM, COMPUTER_LAB, HARDWARE_LAB, SEATER_120, SEATER_240)

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

## Configuration

### Time Settings
Edit the constants at the top of the script:
```python
START_TIME = time(9, 0)          # Day start time
END_TIME = time(18, 30)          # Day end time
LECTURE_DURATION = 3             # 1.5 hours (in 30-min slots)
LAB_DURATION = 4                 # 2 hours (in 30-min slots)
TUTORIAL_DURATION = 2            # 1 hour (in 30-min slots)
SELF_STUDY_DURATION = 2          # 1 hour (in 30-min slots)
```

### Break Times
```python
LUNCH_WINDOW_START = time(12, 30)  # Lunch can start from
LUNCH_WINDOW_END = time(14, 0)     # Lunch must end by
LUNCH_DURATION = 60                # Lunch duration in minutes
```

## Usage

### Basic Usage
Run the script directly:
```bash
python timetable_generator.py
```

This will:
1. Generate `timetable_all_departments.xlsx` - Combined timetable for all departments
2. Generate `all_faculty_timetables.xlsx` - Individual schedules for all faculty
3. Generate `unscheduled_courses.xlsx` - Report of any courses that couldn't be scheduled

### Functions

#### Generate Department Timetables
```python
generate_all_timetables()
```
Creates a single Excel file with separate sheets for each department-semester combination.

#### Check Unscheduled Courses
```python
check_unscheduled_courses()
```
Analyzes the generated timetable and reports courses that weren't fully scheduled according to their L-T-P-S requirements.

#### Generate Faculty Timetables
```python
generate_faculty_timetables()
```
Creates comprehensive timetables for each faculty member showing all their teaching assignments.

## Output Files

### 1. `timetable_all_departments.xlsx`
- **Overview Sheet**: Index of all department-semester combinations
- **Department Sheets**: One sheet per department-semester with:
  - Color-coded timetable grid (Days × Time slots)
  - Course details (code, type, room, faculty)
  - Legend showing course colors and details
  - Self-study only courses list
  - Unscheduled components (if any)

### 2. `all_faculty_timetables.xlsx`
- **Overview Sheet**: List of all faculty with total class count
- **Faculty Sheets**: Individual schedule for each faculty showing:
  - Day-wise schedule
  - Course details
  - Room assignments
  - Department-semester information

### 3. `unscheduled_courses.xlsx`
Report containing:
- Courses that couldn't be fully scheduled
- Required vs. scheduled hours (L-T-P-S)
- Missing components
- Possible reasons for scheduling failure

## Key Features Explained

### Scheduling Constraints

1. **Faculty Constraints**:
   - Maximum 2 course components per day
   - Minimum 3-hour gap between classes (except lab after lecture/tutorial)
   - No scheduling conflicts

2. **Room Allocation**:
   - Automatic selection based on student count
   - Adjacent lab rooms for large lab batches
   - Priority room booking for lectures vs. labs

3. **Break Management**:
   - Morning break: 10:30-10:45 (15 minutes)
   - Staggered lunch breaks for different semesters
   - Inter-class breaks maintained

4. **Basket Courses**:
   - Parallel scheduling of elective courses
   - Shared time slots across basket groups
   - Individual room allocation per course

### Priority Scheduling

Courses are scheduled in priority order:
1. Regular course labs (highest priority)
2. Lectures with high credit hours
3. Tutorials
4. Basket course components (lowest priority)

## Troubleshooting

### Common Issues

**Issue**: "No suitable room found"
- **Solution**: Increase room count in `rooms.csv` or adjust room capacities

**Issue**: "Faculty overbooked"
- **Solution**: Add more faculty options using `/` separator or reduce course load

**Issue**: "Lab sessions not scheduled"
- **Solution**: Ensure sufficient lab rooms of correct type (COMPUTER_LAB/HARDWARE_LAB)

**Issue**: "Courses marked as unscheduled"
- **Solution**: Check `unscheduled_courses.xlsx` for detailed reasons and adjust:
  - Add more rooms
  - Distribute faculty load
  - Adjust semester course count

## Advanced Configuration

### Custom Color Palette
Modify the `COLOR_PALETTE` list to change course colors:
```python
COLOR_PALETTE = [
    "FF5733", "33FF57", "3357FF", ...
]
```

### Room Type Customization
Add new room types in `rooms.csv` and update the `find_suitable_room()` function.

### Section Management
Student sections are automatically calculated from `total_students` column:
- Default maximum: 70 students per section
- Automatically splits large classes into multiple sections

## Notes

- **Slot Duration**: All time slots are 30 minutes
- **Working Days**: Monday to Friday (5 days)
- **Course Credits**: Lecture credits determine number of lecture sessions
- **Self-Study Only**: Courses with only S hours (no L/T/P) are listed separately
- **Basket Course Format**: Use format `B1-COURSECODE`, `B2-COURSECODE`, etc.

## Support

For issues or questions:
1. Check `unscheduled_courses.xlsx` for scheduling problems
2. Verify input CSV formats match requirements
3. Review console output for warnings and errors

## License

This timetable generator is provided as-is for educational and administrative use.

## Team Name
BumbleBee


## Team Members
- Raskin Verma
- Vishwa D
- Sindhu Talari
- Udit Dadhich
