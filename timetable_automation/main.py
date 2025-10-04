import pandas as pd
import numpy as np

# 1. Define days and slot times as in your format (mess times excluded)
days = ["MON", "TUE", "WED", "THU", "FRI"]
slots = [
    "07:30-09:00", "09:00-10:00", "10:00-10:30", "10:30-10:45",
    "10:45-11:00", "11:00-12:00", "12:00-12:15", "12:15-12:30",
    # Mess break skipped: '12:30-13:15', '13:15-14:00'
    "14:00-14:30", "14:30-15:30", "15:30-15:40",
    "15:40-16:00", "16:00-16:30", "16:30-17:10", "17:10-17:30",
    "17:30-18:30"
]

# 2. Create empty timetable grid
timetable = pd.DataFrame('', index=days, columns=slots)

# 3. Dummy Course List: Replace this loop with actual reading of your courses_list
# Each course should have: course_code, L, T, P, S
courses_list = [
    # Example: Course(department, semester, course_code, course_name, L, T, P, S, C, faculty)
    # Course('CSE', 3, 'MA261', 'Differential Equations', 3, 1, 0, 3, 2, 'Dr. Anand Barangi'),
]

# 4. Place lectures (L), tutorials (T), practicals (P) into available slots, skipping mess time and enforcing no double-booking
faculty_slot_tracker = {course.faculty: {day: set() for day in days} for course in courses_list}
batch_slot_tracker = {course.course_code: {day: set() for day in days} for course in courses_list}

def assign_slots(course, n_blocks, label):
    count = 0
    for day in days:
        for slot in slots:
            # Skip mess time already, so no need to check
            if timetable.at[day, slot] == '':
                # Check for >2 consecutive classes (weak version: only for that faculty and batch here)
                if (len(faculty_slot_tracker[course.faculty][day]) < 2 and
                    len(batch_slot_tracker[course.course_code][day]) < 2):
                    timetable.at[day, slot] = f"{course.course_code}-{label}"
                    faculty_slot_tracker[course.faculty][day].add(slot)
                    batch_slot_tracker[course.course_code][day].add(slot)
                    count += 1
                    if count == n_blocks:
                        return
    if count < n_blocks:
        print("Warning: Could not assign all slots for", course.course_code, label)

for course in courses_list:
    assign_slots(course, course.L,  "L")
    assign_slots(course, course.T,  "T")
    assign_slots(course, course.P,  "P")
    # S/self study need not be scheduled

# 5. Export to Excel (Optional: Use openpyxl to color code after export)
timetable.to_excel('data/final_timetable.xlsx', engine='openpyxl')
print("Timetable saved as data/final_timetable.xlsx")

print(timetable)
