import pandas as pd
from timetable_automation.models.batches import Batch
from timetable_automation.models.course import Course
from timetable_automation.models.faculty import Faculty
from timetable_automation.models.room import Room
from timetable_automation.models.timeslots import TimeSlot

# Load CSV data
batches_df = pd.read_csv('data/batches.csv')
courses_df = pd.read_csv('data/courses.csv')
faculty_df = pd.read_csv('data/Faculty.csv')
rooms_df = pd.read_csv('data/rooms.csv')
timeslots_df = pd.read_csv('data/timeslots.csv')

# Create class instances
batches_list = [
    Batch(row['Department'], row['Semester'], row['Total_Students'], row['MaxBatchSize'])
    for _, row in batches_df.iterrows()
]

courses_list = [
    Course(row['Department'], row['Semester'], row['Course Code'], row['Course Name'], row['L-T-P-S-C'], row['Faculty'])
    for _, row in courses_df.iterrows()
]

faculty_list = [
    Faculty(row['Faculty ID'], row['Name'])
    for _, row in faculty_df.iterrows()
]

rooms_list = [
    Room(row['Room ID'], row['Capacity'], row['Type'], row['Facilities'])
    for _, row in rooms_df.iterrows()
]

timeslots_list = [
    TimeSlot(row['Slot_ID'], row['Day'], row['Start_Time'], row['End_Time'])
    for _, row in timeslots_df.iterrows()
]

# Availability trackers
faculty_availability = {f.faculty_id: set() for f in faculty_list}
room_availability = {r.room_id: set() for r in rooms_list}
batch_availability = {b.department: set() for b in batches_list}

def is_slot_free(slot_id, faculty_id, room_id, batch_department):
    return (slot_id not in faculty_availability[faculty_id]
            and slot_id not in room_availability[room_id]
            and slot_id not in batch_availability[batch_department])

def book_slot(slot_id, faculty_id, room_id, batch_department):
    faculty_availability[faculty_id].add(slot_id)
    room_availability[room_id].add(slot_id)
    batch_availability[batch_department].add(slot_id)

# Greedy scheduling
schedule = []
for course in courses_list:
    assigned = False
    # Find matching batch object by department
    batch = next((b for b in batches_list if b.department == course.department and b.semester == course.semester), None)
    if batch is None:
        print(f"No batch found for {course.department} Semester {course.semester}")
        continue
    batch_size = batch.max_batch_size
    # Find faculty by name
    faculty_obj = next((f for f in faculty_list if f.name == course.faculty), None)
    if faculty_obj is None:
        print(f"No faculty found for {course.faculty}")
        continue
    faculty_id = faculty_obj.faculty_id
    for slot in timeslots_list:
        for room in rooms_list:
            # Room capacity and type check (Lab/Class/Other), you may refine as needed
            if room.capacity == '-' or room.capacity == '':
                continue
            if int(room.capacity) >= int(batch_size):
                # Room type match (optional, modify logic as needed)
                if is_slot_free(slot.slot_id, faculty_id, room.room_id, batch.department):
                    book_slot(slot.slot_id, faculty_id, room.room_id, batch.department)
                    schedule.append({
                        'Department': course.department,
                        'Semester': course.semester,
                        'Course Code': course.course_code,
                        'Course Name': course.course_name,
                        'L-T-P-S-C': course.ltp_sc,
                        'Faculty': course.faculty,
                        'Faculty ID': faculty_id,
                        'Slot_ID': slot.slot_id,
                        'Day': slot.day,
                        'Start_Time': slot.start_time,
                        'End_Time': slot.end_time,
                        'Room ID': room.room_id,
                        'Room Type': room.rtype,
                    })
                    assigned = True
                    break
        if assigned:
            break
    if not assigned:
        print(f"Could not assign slot for {course.course_code} ({course.course_name}) in {course.department} Sem {course.semester}")

# Export schedule as CSV
schedule_df = pd.DataFrame(schedule)
schedule_df.to_csv('data/generated_timetable.csv', index=False)
print(schedule_df.head())

