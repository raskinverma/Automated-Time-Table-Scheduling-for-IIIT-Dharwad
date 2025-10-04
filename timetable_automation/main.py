import pandas as pd
from timetable_automation.models.course import Course
from timetable_automation.models.faculty import Faculty
from timetable_automation.models.room import Room
from timetable_automation.models.batches import Batch
from timetable_automation.models.timeslots import TimeSlot

# 1. Load CSVs
courses_df = pd.read_csv('data/courses.csv')
faculty_df = pd.read_csv('data/faculty.csv')
rooms_df = pd.read_csv('data/rooms.csv')
batches_df = pd.read_csv('data/batches.csv')
timeslots_df = pd.read_csv('data/timeslots.csv')

# 2. Build objects
courses_list = [
    Course(row['Department'], row['Semester'], row['Course Code'], row['Course Name'],
           row['L'], row['T'], row['P'], row['S'], row['C'], row['Faculty'])
    for _, row in courses_df.iterrows()
]
faculty_list = [Faculty(row['Faculty ID'], row['Name']) for _, row in faculty_df.iterrows()]
rooms_list = [Room(row['Room ID'], row['Capacity'], row['Type'], row['Facilities']) for _, row in rooms_df.iterrows()]
batches_list = [Batch(row['Department'], row['Semester'], row['Total_Students'], row['MaxBatchSize']) for _, row in batches_df.iterrows()]
slots_list = [TimeSlot(row['Slot_ID'], row['Day'], row['Start_Time'], row['End_Time']) for _, row in timeslots_df.iterrows()]

# 3. Timetable grid and slot mapping
days = ["MON", "TUE", "WED", "THU", "FRI"]
mess_start, mess_end = "12:30", "14:00"
slots_for_grid = []
slot_day_mapping = {}

for _, row in timeslots_df.iterrows():
    st, et = row['Start_Time'].strip(), row['End_Time'].strip()
    if not (et > mess_start and st < mess_end):
        slot_label = f"{st}-{et}"
        slots_for_grid.append(slot_label)
        slot_day_mapping.setdefault(row['Day'], []).append(slot_label)

# Remove duplicates in slot labels
slots_for_grid = list(dict.fromkeys(slots_for_grid))

timetable = pd.DataFrame('', index=days, columns=slots_for_grid)
timetable = timetable.loc[~timetable.index.duplicated(keep='first')]
timetable = timetable.loc[:, ~timetable.columns.duplicated(keep='first')]
assert timetable.index.is_unique, "Timetable index (days) contains duplicates!"
assert timetable.columns.is_unique, "Timetable columns (slots) contains duplicates!"

# 4. Safe assignment function
def assign_slots(course, n_blocks, label):
    count = 0
    for day in days:
        if day not in slot_day_mapping:
            continue
        for slot in slot_day_mapping[day]:
            slot = slot.strip()
            if slot in timetable.columns and day in timetable.index:
                val = timetable.at[day, slot]
                if val == '':
                    timetable.at[day, slot] = f"{course.course_code}-{label}"
                    count += 1
                    if count == int(n_blocks):
                        return
    if count < int(n_blocks):
        print(f"Warning: Could not assign all slots for {course.course_code} {label}")

for course in courses_list:
    assign_slots(course, int(course.L), "L")
    assign_slots(course, int(course.T), "T")
    assign_slots(course, int(course.P), "P")

timetable.to_excel('data/final_timetable.xlsx', engine='openpyxl')
print("Saved timetable to data/final_timetable.xlsx")
print(timetable)
