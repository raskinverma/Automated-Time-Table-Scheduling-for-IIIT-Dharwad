import pandas as pd
import random
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side


# ---------------- Course Class ----------------
class Course:
    def __init__(self, row):
        self.code = str(row["Course_Code"]).strip()
        self.faculty = str(row.get("Faculty", "")).strip()
        self.ltp = str(row["L-T-P-S-C"]).strip()
        self.semester_half = str(row.get("Semester_Half", "0")).strip()
        self.is_elective = str(row.get("Elective", 0)).strip() == "1"

        try:
            self.L, self.T, self.P, self.S, self.C = map(int, self.ltp.split("-"))
        except:
            self.L, self.T, self.P = 0, 0, 0


# ---------------- Scheduler Class ----------------
class Scheduler:
    def __init__(self, slots_file, courses_file, rooms_file):
        # Load time slots
        df = pd.read_csv(slots_file)
        self.slots = [f"{row['Start_Time'].strip()}-{row['End_Time'].strip()}" for _, row in df.iterrows()]
        self.slot_durations = {s: self._slot_duration(s) for s in self.slots}

        # Load courses
        self.courses = [Course(row) for _, row in pd.read_csv(courses_file).iterrows()]

        # Load rooms
        rooms_df = pd.read_csv(rooms_file)
        self.classrooms = rooms_df[rooms_df["Type"].str.lower() == "classroom"]["Room_ID"].tolist()
        self.labs = rooms_df[rooms_df["Type"].str.lower() == "lab"]["Room_ID"].tolist()

        # Constants
        self.days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        self.excluded_slots = ["07:30-09:00", "13:15-14:00", "17:30-18:30"]
        self.MAX_ATTEMPTS = 10

        self.course_room_map = {}

    def _slot_duration(self, slot):
        start, end = slot.split("-")
        h1, m1 = map(int, start.split(":"))
        h2, m2 = map(int, end.split(":"))
        return (h2 + m2 / 60) - (h1 + m1 / 60)

    def _get_free_blocks(self, timetable, day):
        free_blocks, block = [], []
        for slot in self.slots:
            if timetable.at[day, slot] == "" and slot not in self.excluded_slots:
                block.append(slot)
            else:
                if block:
                    free_blocks.append(block)
                    block = []
        if block:
            free_blocks.append(block)
        return free_blocks

    def _allocate_session(self, timetable, lecturer_busy, labs_scheduled, day, faculty, code, duration_hours, session_type="L", is_elective=False):
        # Check if this is a lab session and if a lab is already scheduled for this day
        if session_type == "P" and labs_scheduled[day]:
            return False
        
        free_blocks = self._get_free_blocks(timetable, day)
        for block in free_blocks:
            total = sum(self.slot_durations[s] for s in block)
            if total >= duration_hours:
                slots_to_use, dur_accum = [], 0
                for s in block:
                    slots_to_use.append(s)
                    dur_accum += self.slot_durations[s]
                    if dur_accum >= duration_hours:
                        break

                if not is_elective:
                    if code in self.course_room_map:
                        room = self.course_room_map[code]
                    else:
                        if session_type == "P":
                            if not self.labs:
                                print(f"No labs available for {code}")
                                return False
                            room = random.choice(self.labs)
                        else:
                            if not self.classrooms:
                                print(f"No classrooms available for {code}")
                                return False
                            room = random.choice(self.classrooms)
                        self.course_room_map[code] = room

                for i, s in enumerate(slots_to_use):
                    if session_type == "L":
                        timetable.at[day, s] = f"{code} ({room})" if not is_elective else code
                    elif session_type == "T":
                        timetable.at[day, s] = f"{code}T ({room})" if not is_elective else f"{code}T"
                    elif session_type == "P":
                        timetable.at[day, s] = f"{code} (Lab-{room})" if not is_elective else code

                    # Fill gap slots (like 15-min breaks)
                    if i < len(slots_to_use) - 1:
                        idx = self.slots.index(s)
                        if idx + 1 < len(self.slots):
                            gap_slot = self.slots[idx + 1]
                            if timetable.at[day, gap_slot] == "" and self.slot_durations[gap_slot] == 0.25:
                                timetable.at[day, gap_slot] = "FREE"

                if faculty:
                    lecturer_busy[day].append(faculty)
                
                # Mark that a lab has been scheduled for this day
                if session_type == "P":
                    labs_scheduled[day] = True
                    
                return True
        return False

    def generate_timetable(self, courses_to_allocate, writer, sheet_name):
        timetable = pd.DataFrame("", index=self.days, columns=self.slots)
        lecturer_busy = {day: [] for day in self.days}
        labs_scheduled = {day: False for day in self.days}  # Track lab sessions per day
        self.course_room_map = {}

        electives = [c for c in courses_to_allocate if c.is_elective]
        non_electives = [c for c in courses_to_allocate if not c.is_elective]

        # Pick one elective placeholder
        if electives:
            chosen = random.choice(electives)
            elective_course = Course({
                "Course_Code": "Elective",
                "Faculty": chosen.faculty,
                "L-T-P-S-C": chosen.ltp,
                "Semester_Half": chosen.semester_half,
                "Elective": 0
            })
            non_electives.append(elective_course)

        # Shuffle courses to avoid bias in scheduling order
        random.shuffle(non_electives)
        
        for course in non_electives:
            faculty, code, is_elective = course.faculty, course.code, course.code == "Elective"

            # Lectures - try days in random order for better distribution
            remaining, attempts = course.L, 0
            while remaining > 0 and attempts < self.MAX_ATTEMPTS:
                attempts += 1
                days_to_try = self.days.copy()
                random.shuffle(days_to_try)
                for day in days_to_try:
                    if remaining <= 0 or (faculty and faculty in lecturer_busy[day]):
                        continue
                    alloc = min(1.5, remaining)
                    if self._allocate_session(timetable, lecturer_busy, labs_scheduled, day, faculty, code, alloc, "L", is_elective):
                        remaining -= alloc
                        break

            # Tutorials - try days in random order
            remaining, attempts = course.T, 0
            while remaining > 0 and attempts < self.MAX_ATTEMPTS:
                attempts += 1
                days_to_try = self.days.copy()
                random.shuffle(days_to_try)
                for day in days_to_try:
                    if remaining <= 0 or (faculty and faculty in lecturer_busy[day]):
                        continue
                    if self._allocate_session(timetable, lecturer_busy, labs_scheduled, day, faculty, code, 1, "T", is_elective):
                        remaining -= 1
                        break

            # Practicals - prioritize days without labs first
            remaining, attempts = course.P, 0
            while remaining > 0 and attempts < self.MAX_ATTEMPTS:
                attempts += 1
                # Prioritize days without labs
                days_without_labs = [d for d in self.days if not labs_scheduled[d]]
                days_to_try = days_without_labs.copy()
                random.shuffle(days_to_try)
                
                for day in days_to_try:
                    if remaining <= 0 or (faculty and faculty in lecturer_busy[day]):
                        continue
                    alloc = min(2, remaining) if remaining >= 2 else remaining
                    if self._allocate_session(timetable, lecturer_busy, labs_scheduled, day, faculty, code, alloc, "P", is_elective):
                        remaining -= alloc
                        break

        # Clear excluded slots
        for day in self.days:
            for slot in self.excluded_slots:
                if slot in timetable.columns:
                    timetable.at[day, slot] = ""

        timetable.to_excel(writer, sheet_name=sheet_name, index=True)
        print(f"Saved timetable to sheet '{sheet_name}'")

    def run(self, output_file="timetable_full.xlsx"):
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            self.generate_timetable([c for c in self.courses if c.semester_half in ["1", "0"]],
                                    writer, "First_Half")
            self.generate_timetable([c for c in self.courses if c.semester_half in ["2", "0"]],
                                    writer, "Second_Half")

        # Remove default sheet (Sheet/Sheet1)
        wb = load_workbook(output_file)
        for default in ["Sheet", "Sheet1"]:
            if default in wb.sheetnames:
                wb.remove(wb[default])
        wb.save(output_file)

        self.format_excel(output_file)

    def format_excel(self, filename):
        wb = load_workbook(filename)
        thin_border = Border(left=Side(style='thin'),
                             right=Side(style='thin'),
                             top=Side(style='thin'),
                             bottom=Side(style='thin'))

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            for row in range(2, ws.max_row + 1):
                start_col = 2
                while start_col <= ws.max_column:
                    cell = ws.cell(row=row, column=start_col)
                    if cell.value and cell.value != "FREE":
                        merge_count = 0
                        for col in range(start_col + 1, ws.max_column + 1):
                            if ws.cell(row=row, column=col).value == cell.value:
                                merge_count += 1
                            else:
                                break
                        if merge_count > 0:
                            ws.merge_cells(start_row=row, start_column=start_col,
                                           end_row=row, end_column=start_col + merge_count)
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        for col_idx in range(start_col, start_col + merge_count + 1):
                            ws.cell(row=row, column=col_idx).border = thin_border
                        start_col += merge_count + 1
                    else:
                        cell.border = thin_border
                        start_col += 1

            # Format headers
            for col in range(1, ws.max_column + 1):
                ws.cell(row=1, column=col).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(row=1, column=col).border = thin_border

        wb.save(filename)
        print(f"Formatted timetable with borders saved in {filename}")


# ---------------- Run ----------------
if __name__ == "__main__":
    scheduler = Scheduler("data/timeslots.csv", "data/courses.csv", "data/rooms.csv")
    scheduler.run("timetable_full.xlsx")