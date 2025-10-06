import pandas as pd
import random
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

RANDOM_SEED = 42
random.seed(RANDOM_SEED)


class Course:
    def __init__(self, row):
        self.code = str(row["Course_Code"]).strip()
        self.faculty = str(row.get("Faculty", "")).strip()
        self.ltp = str(row["L-T-P-S-C"]).strip()
        self.semester_half = str(row.get("Semester_Half", "0")).strip()
        self.is_elective = str(row.get("Elective", 0)).strip() == "1"
        self.class_name = str(row.get("Class", "")).strip()

        try:
            self.L, self.T, self.P, self.S, self.C = map(int, self.ltp.split("-"))
        except:
            self.L, self.T, self.P = 0, 0, 0


class Scheduler:
    def __init__(self, slots_file, courses_file, rooms_file, global_room_usage):
        df = pd.read_csv(slots_file)
        self.slots = [f"{row['Start_Time'].strip()}-{row['End_Time'].strip()}" for _, row in df.iterrows()]
        self.slot_durations = {s: self._slot_duration(s) for s in self.slots}

        self.courses = [Course(row) for _, row in pd.read_csv(courses_file).iterrows()]

        rooms_df = pd.read_csv(rooms_file)
        self.classrooms = rooms_df[rooms_df["Type"].str.lower() == "classroom"]["Room_ID"].tolist()
        self.labs = rooms_df[rooms_df["Type"].str.lower() == "lab"]["Room_ID"].tolist()

        self.days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        self.excluded_slots = ["07:30-09:00", "13:15-14:00", "17:30-18:30"]
        self.MAX_ATTEMPTS = 10
        self.course_room_map = {}
        self.global_room_usage = global_room_usage

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

    def _allocate_session(
        self, timetable, lecturer_busy, labs_scheduled, day, faculty, code, duration_hours, session_type="L", is_elective=False
    ):
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
                        possible_rooms = self.labs if session_type == "P" else self.classrooms
                        available_rooms = [
                            r
                            for r in possible_rooms
                            if all(
                                r not in self.global_room_usage.get(day, {}).get(s, [])
                                for s in slots_to_use
                            )
                        ]
                        if not available_rooms:
                            return False
                        room = random.choice(available_rooms)
                        self.course_room_map[code] = room

                    for s in slots_to_use:
                        self.global_room_usage.setdefault(day, {}).setdefault(s, []).append(room)
                else:
                    room = ""

                for i, s in enumerate(slots_to_use):
                    if session_type == "L":
                        timetable.at[day, s] = f"{code} ({room})" if not is_elective else "Elective"
                    elif session_type == "T":
                        timetable.at[day, s] = f"{code}T ({room})" if not is_elective else "Elective"
                    elif session_type == "P":
                        timetable.at[day, s] = f"{code} (Lab-{room})" if not is_elective else "Elective"

                    if i < len(slots_to_use) - 1:
                        idx = self.slots.index(s)
                        if idx + 1 < len(self.slots):
                            gap_slot = self.slots[idx + 1]
                            if timetable.at[day, gap_slot] == "" and self.slot_durations[gap_slot] == 0.25:
                                timetable.at[day, gap_slot] = "FREE"

                if faculty:
                    lecturer_busy[day].append(faculty)

                if session_type == "P":
                    labs_scheduled[day] = True

                return True
        return False

    def generate_timetable(self, courses_to_allocate, writer, sheet_name):
        timetable = pd.DataFrame("", index=self.days, columns=self.slots)
        lecturer_busy = {day: [] for day in self.days}
        labs_scheduled = {day: False for day in self.days}
        self.course_room_map = {}

        electives = [c for c in courses_to_allocate if c.is_elective]
        non_electives = [c for c in courses_to_allocate if not c.is_elective]

        elective_info = []
        for c in electives:
            elective_info.append({
                "code": c.code,
                "faculty": c.faculty,
                "class": c.class_name
            })

        if electives:
            chosen = random.choice(electives)
            elective_course = Course(
                {
                    "Course_Code": "Elective",
                    "Faculty": chosen.faculty,
                    "L-T-P-S-C": chosen.ltp,
                    "Semester_Half": chosen.semester_half,
                    "Elective": 0,
                    "Class": chosen.class_name
                }
            )
            non_electives.append(elective_course)

        random.shuffle(non_electives)

        for course in non_electives:
            faculty, code, is_elective = course.faculty, course.code, course.code == "Elective"

            remaining, attempts = course.L, 0
            while remaining > 0 and attempts < self.MAX_ATTEMPTS:
                attempts += 1
                days_to_try = self.days.copy()
                random.shuffle(days_to_try)
                for day in days_to_try:
                    if remaining <= 0 or (faculty and faculty in lecturer_busy[day]):
                        continue
                    alloc = min(1.5, remaining)
                    if self._allocate_session(
                        timetable, lecturer_busy, labs_scheduled, day, faculty, code, alloc, "L", is_elective
                    ):
                        remaining -= alloc
                        break

            remaining, attempts = course.T, 0
            while remaining > 0 and attempts < self.MAX_ATTEMPTS:
                attempts += 1
                days_to_try = self.days.copy()
                random.shuffle(days_to_try)
                for day in days_to_try:
                    if remaining <= 0 or (faculty and faculty in lecturer_busy[day]):
                        continue
                    if self._allocate_session(
                        timetable, lecturer_busy, labs_scheduled, day, faculty, code, 1, "T", is_elective
                    ):
                        remaining -= 1
                        break

            remaining, attempts = course.P, 0
            while remaining > 0 and attempts < self.MAX_ATTEMPTS:
                attempts += 1
                days_without_labs = [d for d in self.days if not labs_scheduled[d]]
                days_to_try = days_without_labs.copy()
                random.shuffle(days_to_try)

                for day in days_to_try:
                    if remaining <= 0 or (faculty and faculty in lecturer_busy[day]):
                        continue
                    alloc = min(2, remaining) if remaining >= 2 else remaining
                    if self._allocate_session(
                        timetable, lecturer_busy, labs_scheduled, day, faculty, code, alloc, "P", is_elective
                    ):
                        remaining -= alloc
                        break

        for day in self.days:
            for slot in self.excluded_slots:
                if slot in timetable.columns:
                    timetable.at[day, slot] = ""

        timetable.to_excel(writer, sheet_name=sheet_name, index=True)
        print(f"Saved timetable to sheet '{sheet_name}'")
        return elective_info

    def run(self, output_file="timetable_full.xlsx"):
        all_electives_info = {}
        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            all_electives_info["First_Half"] = self.generate_timetable(
                [c for c in self.courses if c.semester_half in ["1", "0"]], writer, "First_Half"
            )
            all_electives_info["Second_Half"] = self.generate_timetable(
                [c for c in self.courses if c.semester_half in ["2", "0"]], writer, "Second_Half"
            )

        wb = load_workbook(output_file)
        for default in ["Sheet", "Sheet1"]:
            if default in wb.sheetnames and len(wb.sheetnames) > 1:
                wb.remove(wb[default])
        wb.save(output_file)
        self.format_excel(output_file, all_electives_info)

    def format_excel(self, filename, all_electives_info):
        wb = load_workbook(filename)
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        color_map = {}
        palette = ["FFC7CE", "C6EFCE", "FFEB9C", "BDD7EE", "D9EAD3", "F4CCCC", "D9D2E9", "FCE5CD", "C9DAF8", "EAD1DC"]
        color_index = 0
        elective_color = "FFD966"

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            for row in range(2, ws.max_row + 1):
                start_col = 2
                while start_col <= ws.max_column:
                    cell = ws.cell(row=row, column=start_col)
                    if cell.value and cell.value != "FREE":
                        raw_code = cell.value.split(" ")[0]
                        code = raw_code.rstrip("T")
                        if code != "Elective" and code not in color_map:
                            color_map[code] = palette[color_index % len(palette)]
                            color_index += 1
                        fill = PatternFill(
                            start_color=color_map.get(code, elective_color),
                            end_color=color_map.get(code, elective_color),
                            fill_type="solid"
                        )
                        cell.fill = fill
                        merge_count = 0
                        for col in range(start_col + 1, ws.max_column + 1):
                            next_cell = ws.cell(row=row, column=col)
                            if next_cell.value == cell.value:
                                next_cell.fill = fill
                                merge_count += 1
                            else:
                                break
                        if merge_count > 0:
                            ws.merge_cells(
                                start_row=row, start_column=start_col, end_row=row, end_column=start_col + merge_count
                            )
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        for col_idx in range(start_col, start_col + merge_count + 1):
                            ws.cell(row=row, column=col_idx).border = thin_border
                        start_col += merge_count + 1
                    else:
                        cell.border = thin_border
                        start_col += 1

            # Legend
            start_row = ws.max_row + 3
            ws.cell(start_row, 2, "Course Code").border = thin_border
            ws.cell(start_row, 3, "Faculty").border = thin_border
            ws.cell(start_row, 4, "Color").border = thin_border
            ws.cell(start_row, 2).alignment = ws.cell(start_row, 3).alignment = ws.cell(
                start_row, 4
            ).alignment = Alignment(horizontal="center", vertical="center")

            for i, code in enumerate(color_map, start=1):
                ws.cell(start_row + i, 2, code).border = thin_border
                faculty = next((c.faculty for c in self.courses if c.code == code), "")
                ws.cell(start_row + i, 3, faculty).border = thin_border
                ws.cell(start_row + i, 4, "").fill = PatternFill(start_color=color_map[code], end_color=color_map[code], fill_type="solid")
                ws.cell(start_row + i, 4).border = thin_border

            # Grouped Electives
            ws.cell(start_row + len(color_map) + 2, 2, "Electives").border = thin_border
            ws.cell(start_row + len(color_map) + 2, 4, "").fill = PatternFill(start_color=elective_color, end_color=elective_color, fill_type="solid")
            ws.cell(start_row + len(color_map) + 2, 2).alignment = Alignment(horizontal="center")
            ws.cell(start_row + len(color_map) + 2, 4).border = thin_border

            elective_info = all_electives_info.get(sheet_name, [])
            for j, info in enumerate(elective_info, start=1):
                ws.cell(start_row + len(color_map) + 2 + j, 2, info["code"]).border = thin_border
                ws.cell(start_row + len(color_map) + 2 + j, 3, info["faculty"]).border = thin_border
                ws.cell(start_row + len(color_map) + 2 + j, 4, info["class"]).border = thin_border

            # ---------- AUTO COLUMN WIDTH ----------
            for col in ws.columns:
                max_length = 0
                column = col[0].column
                column_letter = get_column_letter(column)
                for cell in col:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                ws.column_dimensions[column_letter].width = max_length + 2

        wb.save(filename)
        print(f"Formatted timetable with borders, colors, grouped electives, and auto column widths saved in {filename}")


if __name__ == "__main__":
    departments = {
        "CSE": "data/CSE_courses.csv",
        "DSAI": "data/DSAI_courses.csv",
    }
    rooms_file = "data/rooms.csv"
    slots_file = "data/timeslots.csv"

    global_room_usage = {}

    for dept_name, course_file in departments.items():
        print(f"\nGenerating timetable for {dept_name}...")
        scheduler = Scheduler(slots_file, course_file, rooms_file, global_room_usage)
        scheduler.run(f"{dept_name}_timetable.xlsx")
