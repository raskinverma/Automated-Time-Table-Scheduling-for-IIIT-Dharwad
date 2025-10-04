import pandas as pd
from timetable_automation.models.course import Course
from timetable_automation.models.faculty import Faculty
from timetable_automation.models.room import Room
from timetable_automation.models.batches import Batch
from timetable_automation.models.timeslots import TimeSlot

# 1. Extract data from all CSVs
courses_df = pd.read_csv('data/courses.csv')
faculty_df = pd.read_csv('data/faculty.csv')
rooms_df = pd.read_csv('data/rooms.csv')
batches_df = pd.read_csv('data/batches.csv')
timeslots_df = pd.read_csv('data/timeslots.csv')

courses_list = [
    Course(row['Department'], row['Semester'], row['Course Code'], row['Course Name'],
           row['L'], row['T'], row['P'], row['S'], row['C'], row['Faculty'])
    for _, row in courses_df.iterrows()
]
faculty_list = [Faculty(row['Faculty ID'], row['Name']) for _, row in faculty_df.iterrows()]
rooms_list   = [Room(row['Room ID'], row['Capacity'], row['Type'], row['Facilities']) for _, row in rooms_df.iterrows()]
batches_list = [Batch(row['Department'], row['Semester'], row['Total_Students'], row['MaxBatchSize']) for _, row in batches_df.iterrows()]
slots_list   = [TimeSlot(row['Slot_ID'], row['Day'], row['Start_Time'], row['End_Time']) for _, row in timeslots_df.iterrows()]

# 2. Timetable Grid, skip mess slots (12:30-14:00)
days = ["MON", "TUE", "WED", "THU", "FRI"]
slots_for_grid = []
for _, row in timeslots_df.iterrows():
    # Exclude any slot overlapping with mess time (12:30-14:00)
    if not (row['End_Time'] > "12:30" and row['Start_Time'] < "14:00"):
        slots_for_grid.append(f"{row['Start_Time']}-{row['End_Time']}")

timetable = pd.DataFrame('', index=days, columns=slots_for_grid)

# 3. Build per-day slot record (to preserve day correspondence)
slots_by_day = {day: [slot for slot in slots_for_grid if timeslots_df[(timeslots_df['Day']==day) &
                    (timeslots_df['Start_Time']==slot.split('-')[0])].shape[0] > 0] for day in days}

# 4. Greedy block assignment (can be improved per faculty, room, overlap, etc)
def assign_slots(course, n_blocks, label):
    count = 0
    for day in days:
        for slot in slots_by_day[day]:
            if timetable.at[day, slot] == '':
                timetable.at[day, slot] = f"{course.course_code}-{label}"
                count += 1
                if count == int(n_blocks):
                    return
    if count < n_blocks:
        print(f"Warning: Could not assign all slots for {course.course_code} {label}")

for course in courses_list:
    assign_slots(course, int(course.L), "L")
    assign_slots(course, int(course.T), "T")
    assign_slots(course, int(course.P), "P")
    # S (self-study) not scheduled on grid

# 5. Export timetable for grid view styling
timetable.to_excel('data/final_timetable.xlsx', engine='openpyxl')
print("Saved timetable to data/final_timetable.xlsx")
print(timetable)
