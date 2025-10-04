import pandas as pd
import random

# ---------------- Load Time Slots ----------------
df = pd.read_csv('data/timeslots.csv')

slots = [{"start": row["Start_Time"], "end": row["End_Time"]} for _, row in df.iterrows()]
slot_keys = [f"{slot['start'].strip()}-{slot['end'].strip()}" for slot in slots]

def slot_duration(slot):
    start, end = slot.split("-")
    h1, m1 = map(int, start.split(":"))
    h2, m2 = map(int, end.split(":"))
    return (h2 + m2 / 60) - (h1 + m1 / 60)

slot_durations = {s: slot_duration(s) for s in slot_keys}

# ---------------- Load Courses and Rooms ----------------
courses = pd.read_csv("data/courses.csv").to_dict(orient="records")
rooms_df = pd.read_csv("data/rooms.csv")  # must have columns: Room_ID, Type
classrooms = rooms_df[rooms_df["Type"].str.lower() == "classroom"]["Room_ID"].tolist()
labs = rooms_df[rooms_df["Type"].str.lower() == "lab"]["Room_ID"].tolist()

# ---------------- Constants ----------------
days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
excluded_slots = ["07:30-09:00", "13:15-14:00", "17:30-18:30"]
MAX_ATTEMPTS = 10

# ---------------- Functions ----------------
def get_free_blocks(timetable, day):
    free_blocks = []
    block = []
    for slot in slot_keys:
        if timetable.at[day, slot] == "" and slot not in excluded_slots:
            block.append(slot)
        else:
            if block:
                free_blocks.append(block)
                block = []
    if block:
        free_blocks.append(block)
    return free_blocks

# Track rooms assigned per course to ensure consistency
course_room_map = {}

def allocate_session(timetable, lecturer_busy, day, faculty, code, duration_hours, session_type="L", is_elective=False):
    free_blocks = get_free_blocks(timetable, day)
    for block in free_blocks:
        total = sum(slot_durations[s] for s in block)
        if total >= duration_hours:
            slots_to_use = []
            dur_accum = 0
            for s in block:
                slots_to_use.append(s)
                dur_accum += slot_durations[s]
                if dur_accum >= duration_hours:
                    break

            # Assign consistent room based on session type (skip for electives)
            if not is_elective:
                if code in course_room_map:
                    room = course_room_map[code]
                else:
                    if session_type == "P":
                        if not labs:
                            print(f"No labs available for {code}")
                            return False
                        room = random.choice(labs)
                    else:  # L or T
                        if not classrooms:
                            print(f"No classrooms available for {code}")
                            return False
                        room = random.choice(classrooms)
                    course_room_map[code] = room

            # Allocate slots, adding 15-min gap if available
            for i, s in enumerate(slots_to_use):
                if session_type == "L":
                    timetable.at[day, s] = f"{code} ({room})" if not is_elective else code
                elif session_type == "T":
                    timetable.at[day, s] = f"{code}T ({room})" if not is_elective else f"{code}T"
                elif session_type == "P":
                    timetable.at[day, s] = f"{code} (Lab-{room})" if not is_elective else code

                # Add 15-min gap after the slot if next slot is consecutive and 15-min
                if i < len(slots_to_use) - 1:
                    idx = slot_keys.index(s)
                    next_idx = idx + 1
                    if next_idx < len(slot_keys):
                        gap_slot = slot_keys[next_idx]
                        if timetable.at[day, gap_slot] == "" and slot_durations[gap_slot] == 0.25:
                            timetable.at[day, gap_slot] = "FREE"

            if faculty:
                lecturer_busy[day].append(faculty)
            return True
    return False

def generate_timetable(courses_to_allocate, writer=None, sheet_name="Sheet1"):
    timetable = pd.DataFrame("", index=days, columns=slot_keys)
    lecturer_busy = {day: [] for day in days}
    global course_room_map
    course_room_map = {}

    # ---------------- Separate electives ----------------
    electives = [c for c in courses_to_allocate if str(c.get("Elective", 0)) == "1"]
    non_electives = [c for c in courses_to_allocate if str(c.get("Elective", 0)) != "1"]

    # ---------------- Pick only one elective ----------------
    if electives:
        chosen_elective = random.choice(electives)
        elective_course = {
            "Course_Code": "Elective",
            "Faculty": chosen_elective.get("Faculty", ""),
            "L-T-P-S-C": chosen_elective["L-T-P-S-C"]
        }
        non_electives.append(elective_course)

    # ---------------- Allocate all courses ----------------
    for course in non_electives:
        faculty = str(course.get("Faculty", "")).strip()
        code = str(course["Course_Code"]).strip()
        is_elective = True if code == "Elective" else False

        try:
            L, T, P, S, C = map(int, [x.strip() for x in course["L-T-P-S-C"].split("-")])
        except:
            continue

        # --- Lectures ---
        lecture_hours_remaining = L
        attempts = 0
        while lecture_hours_remaining > 0 and attempts < MAX_ATTEMPTS:
            attempts += 1
            for day in days:
                if lecture_hours_remaining <= 0 or (faculty and faculty in lecturer_busy[day]):
                    continue
                alloc_hours = min(1.5, lecture_hours_remaining)
                if allocate_session(timetable, lecturer_busy, day, faculty, code, alloc_hours, "L", is_elective):
                    lecture_hours_remaining -= alloc_hours
                    break
        if lecture_hours_remaining > 0:
            print(f"Warning: Could not fully allocate lectures for {code}")

        # --- Tutorials ---
        tutorial_hours_remaining = T
        attempts = 0
        while tutorial_hours_remaining > 0 and attempts < MAX_ATTEMPTS:
            attempts += 1
            for day in days:
                if tutorial_hours_remaining <= 0 or (faculty and faculty in lecturer_busy[day]):
                    continue
                if allocate_session(timetable, lecturer_busy, day, faculty, code, 1, "T", is_elective):
                    tutorial_hours_remaining -= 1
                    break
        if tutorial_hours_remaining > 0:
            print(f"Warning: Could not fully allocate tutorials for {code}")

        # --- Practicals ---
        practical_hours_remaining = P
        attempts = 0
        while practical_hours_remaining > 0 and attempts < MAX_ATTEMPTS:
            attempts += 1
            for day in days:
                if practical_hours_remaining <= 0 or (faculty and faculty in lecturer_busy[day]):
                    continue
                alloc_hours = min(2, practical_hours_remaining) if practical_hours_remaining >= 2 else practical_hours_remaining
                if allocate_session(timetable, lecturer_busy, day, faculty, code, alloc_hours, "P", is_elective):
                    practical_hours_remaining -= alloc_hours
                    break
        if practical_hours_remaining > 0:
            print(f"Warning: Could not fully allocate practicals for {code}")

    # Clear excluded slots
    for day in days:
        for slot in excluded_slots:
            if slot in timetable.columns:
                timetable.at[day, slot] = ""

    # Save timetable to Excel sheet
    if writer:
        timetable.to_excel(writer, sheet_name=sheet_name, index=True)
        print(f"Saved timetable to sheet '{sheet_name}'")
    else:
        timetable.to_excel(sheet_name + ".xlsx", index=True)
        print(f"Saved timetable to {sheet_name}.xlsx")


# ---------------- Split Courses by Semester Half ----------------
courses_first_half = [c for c in courses if str(c.get("Semester_Half")).strip() in ["1", "0"]]
courses_second_half = [c for c in courses if str(c.get("Semester_Half")).strip() in ["2", "0"]]

# ---------------- Generate Timetables into one Excel file ----------------
with pd.ExcelWriter("timetable_full.xlsx", engine="openpyxl") as writer:
    generate_timetable(courses_first_half, writer, sheet_name="First_Half")
    generate_timetable(courses_second_half, writer, sheet_name="Second_Half")
