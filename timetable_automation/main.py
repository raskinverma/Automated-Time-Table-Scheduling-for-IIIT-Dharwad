import pandas as pd
from timetable_automation.models.faculty import Faculty
from timetable_automation.models.course import Course
from timetable_automation.models.room import Room
from timetable_automation.models.batch import Batch
from timetable_automation.models.timeslot import TimeSlot

# Load CSV data from data/ directory
faculty_df = pd.read_csv('data/faculty.csv')
courses_df = pd.read_csv('data/courses.csv')
batches_df = pd.read_csv('data/batches.csv')
rooms_df = pd.read_csv('data/rooms.csv')
timeslots_df = pd.read_csv('data/timeslots.csv')

# Instantiate Python objects for each entity
faculty_list = [Faculty(*row) for row in faculty_df.values]
courses_list = [Course(*row) for row in courses_df.values]
batches_list = [Batch(*row) for row in batches_df.values]
rooms_list = [Room(*row) for row in rooms_df.values]
timeslots_list = [TimeSlot(*row) for row in timeslots_df.values]

# Prepare availability trackers
faculty_availability = {f.id: set() for f in faculty_list}
room_availability = {r.name: set() for r in rooms_list}
batch_availability = {b.id: set() for b in batches_list}

def is_slot_free(slot_id, faculty_id, room_name, batch_id):
    if slot_id in faculty_availability[faculty_id]:
        return False
    if slot_id in room_availability[room_name]:
        return False
    if slot_id in batch_availability[batch_id]:
        return False
    return True

def book_slot(slot_id, faculty_id, room_name, batch_id):
    faculty_availability[faculty_id].add(slot_id)
    room_availability[room_name].add(slot_id)
    batch_availability[batch_id].add(slot_id)

# Greedy scheduling
schedule = []

for course in courses_list:
    assigned = False
    batch_size = next(b.size for b in batches_list if b.id == course.batch_id)
    for slot in timeslots_list:
        for room in rooms_list:
            correct_room = (course.ctype == 'Lab' and room.rtype == 'Lab') or (course.ctype != 'Lab' and room.rtype == 'Theory')
            if correct_room and room.capacity >= batch_size:
                if is_slot_free(slot.slot_id, course.faculty_id, room.name, course.batch_id):
                    book_slot(slot.slot_id, course.faculty_id, room.name, course.batch_id)
                    schedule.append({
                        'Course_ID': course.id,
                        'Course_Name': course.name,
                        'Faculty_ID': course.faculty_id,
                        'Batch_ID': course.batch_id,
                        'Slot_ID': slot.slot_id,
                        'Day': slot.day,
                        'Start_Time': slot.start_time,
                        'End_Time': slot.end_time,
                        'Room': room.name,
                    })
                    assigned = True
                    break
        if assigned:
            break
    if not assigned:
        print(f"Could not assign slot for course {course.name} (Batch {course.batch_id})")

# Output schedule as CSV
schedule_df = pd.DataFrame(schedule)
schedule_df.to_csv('data/generated_timetable.csv', index=False)
print(schedule_df.head())
