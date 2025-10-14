import pandas as pd
import random
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill
RANDOM_SEED = 42
random.seed(RANDOM_SEED)

class Course:
    def __init__(self, row):
        self.code = str(row["Course_Code"]).strip()
        self.title = str(row.get("Course_Title", self.code)).strip()
        self.faculty = str(row.get("Faculty", "")).strip()
        self.ltp = str(row["L-T-P-S-C"]).strip()
        self.semester_half = str(row.get("Semester_Half", "0")).strip()
        self.is_elective = str(row.get("Elective", 0)).strip() == "1"
        try:
            self.L, self.T, self.P, self.S, self.C = map(int, self.ltp.split("-"))
        except Exception:
            self.L, self.T, self.P = 0, 0, 0


class Scheduler:
    def __init__(self, slots_file, courses_file, rooms_file, global_room_usage):
        df = pd.read_csv(slots_file)
        self.slots = [f"{row['Start_Time'].strip()}-{row['End_Time'].strip()}" for _, row in df.iterrows()]
        self.slot_durations = {s: self._slot_duration(s) for s in self.slots}

        courses_df = pd.read_csv(courses_file)
        self.courses = [Course(row) for _, row in courses_df.iterrows()]

        rooms_df = pd.read_csv(rooms_file)
        self.classrooms = rooms_df[rooms_df["Type"].str.lower() == "classroom"]["Room_ID"].tolist()
        self.labs = rooms_df[rooms_df["Type"].str.lower() == "lab"]["Room_ID"].tolist()

        self.days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
        self.excluded_slots = ["07:30-09:00", "10:30-10:45", "13:15-14:00", "17:30-18:30"]
        self.MAX_ATTEMPTS = 10

        self.course_room_map = {}      
        self.global_room_usage = global_room_usage 
        self.scheduled_entries = []  
        self.electives_by_sheet = {}  
        self.elective_room_assignment = {}  

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
        self,
        timetable,
        lecturer_busy,
        labs_scheduled,
        day,
        faculty,
        code,
        duration_hours,
        session_type="L",
        is_elective=False,
        sheet_name=None,
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
                            if all(r not in self.global_room_usage.get(day, {}).get(s, []) for s in slots_to_use)
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
                        display_text = f"{code} ({room})" if (room and not is_elective) else code
                    elif session_type == "T":
                        display_text = f"{code}T ({room})" if (room and not is_elective) else f"{code}T"
                    elif session_type == "P":
                        display_text = f"{code} (Lab-{room})" if (room and not is_elective) else code
                    else:
                        display_text = code

                    timetable.at[day, s] = display_text

                    self.scheduled_entries.append(
                        {
                            "sheet": sheet_name,
                            "day": day,
                            "slot": s,
                            "code": code,
                            "display": display_text,
                            "faculty": faculty,
                            "room": room,
                        }
                    )

                    
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

        
        self.electives_by_sheet[sheet_name] = electives

        
        if electives:
            chosen = random.choice(electives)
            elective_course = Course(
                {
                    "Course_Code": "Elective",
                    "Course_Title": chosen.title,
                    "Faculty": chosen.faculty,
                    "L-T-P-S-C": chosen.ltp,
                    "Semester_Half": chosen.semester_half,
                    "Elective": 0,
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
                        timetable, lecturer_busy, labs_scheduled, day, faculty, code, alloc, "L", is_elective, sheet_name
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
                        timetable, lecturer_busy, labs_scheduled, day, faculty, code, 1, "T", is_elective, sheet_name
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
                        timetable, lecturer_busy, labs_scheduled, day, faculty, code, alloc, "P", is_elective, sheet_name
                    ):
                        remaining -= alloc
                        break

       
        for day in self.days:
            for slot in self.excluded_slots:
                if slot in timetable.columns:
                    timetable.at[day, slot] = ""

        timetable.to_excel(writer, sheet_name=sheet_name, index=True)
        print(f"Saved timetable to sheet '{sheet_name}'")

    def _compute_elective_room_assignments_legally(self, sheet_name):
        
        electives = self.electives_by_sheet.get(sheet_name, [])
        if not electives:
            self.elective_room_assignment[sheet_name] = {}
            return

        
        elective_slots = [(ent["day"], ent["slot"]) for ent in self.scheduled_entries if ent["sheet"] == sheet_name and ent["code"] == "Elective"]
       
        elective_slots = sorted(list(dict.fromkeys(elective_slots)))

       
        candidate_rooms = list(self.classrooms) + list(self.labs)
        room_free_all_slots = {}
        for r in candidate_rooms:
            ok = True
            for day, slot in elective_slots:
                if r in self.global_room_usage.get(day, {}).get(slot, []):
                    ok = False
                    break
            room_free_all_slots[r] = ok

      
        free_rooms = [r for r, ok in room_free_all_slots.items() if ok]
       
        if not free_rooms:
            
            room_free_counts = {r: 0 for r in candidate_rooms}
            for r in candidate_rooms:
                for day, slot in elective_slots:
                    if r not in self.global_room_usage.get(day, {}).get(slot, []):
                        room_free_counts[r] += 1
            
            ordered = sorted(candidate_rooms, key=lambda x: (-room_free_counts[x], x))
            free_rooms = ordered

        
        assigned = {}
        used = set()
        idx = 0
        for e in electives:
            
            chosen = None
            for r in free_rooms:
                if r in used:
                    continue
               
                ok = True
                for day, slot in elective_slots:
                    if r in self.global_room_usage.get(day, {}).get(slot, []):
                        ok = False
                        break
                if ok:
                    chosen = r
                    break
            if not chosen:
                
                chosen = free_rooms[idx % len(free_rooms)] if free_rooms else ""
            assigned[e.code + "||" + e.title] = chosen or ""
            used.add(chosen)
            idx += 1

        self.elective_room_assignment[sheet_name] = assigned

    def format_student_timetable_with_legend(self, filename):
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

       
        for sheet_name in wb.sheetnames:
            
            self._compute_elective_room_assignments_legally(sheet_name)

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            for row in range(2, ws.max_row + 1):
                start_col = 2
                while start_col <= ws.max_column:
                    cell = ws.cell(row=row, column=start_col)
                    if cell.value and cell.value != "FREE":
                        raw_code = str(cell.value).split(" ")[0]
                        code = raw_code.rstrip("T")
                        if code not in color_map:
                            color_map[code] = palette[color_index % len(palette)]
                            color_index += 1
                        fill = PatternFill(start_color=color_map[code], end_color=color_map[code], fill_type="solid")
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

           
            start_row = ws.max_row + 3
            ws.cell(start_row, 2, "S.No").border = thin_border
            ws.cell(start_row, 3, "Course Code").border = thin_border
            ws.cell(start_row, 4, "Course Title").border = thin_border
            ws.cell(start_row, 5, "Faculty").border = thin_border
            ws.cell(start_row, 6, "Classroom").border = thin_border
            ws.cell(start_row, 7, "Color").border = thin_border
            for c in range(2, 8):
                ws.cell(start_row, c).alignment = Alignment(horizontal="center", vertical="center")

            
            i = 1
            for code in color_map:
                if code == "Elective":
                    continue
                ws.cell(start_row + i, 2, i).border = thin_border
                ws.cell(start_row + i, 3, code).border = thin_border
                course_name = next((c.title for c in self.courses if c.code == code), code)
                ws.cell(start_row + i, 4, course_name).border = thin_border
                faculty = next((c.faculty for c in self.courses if c.code == code), "")
                ws.cell(start_row + i, 5, faculty).border = thin_border
                classroom = self.course_room_map.get(code, "") or ""
                ws.cell(start_row + i, 6, classroom).border = thin_border
                ws.cell(start_row + i, 7, "").fill = PatternFill(start_color=color_map[code], end_color=color_map[code], fill_type="solid")
                ws.cell(start_row + i, 7).border = thin_border
                i += 1

            
            electives = self.electives_by_sheet.get(sheet_name, [])
            elective_assignment = self.elective_room_assignment.get(sheet_name, {})

            for elective in electives:
                ws.cell(start_row + i, 2, i).border = thin_border
                ws.cell(start_row + i, 3, "Elective").border = thin_border
                ws.cell(start_row + i, 4, elective.title).border = thin_border
                ws.cell(start_row + i, 5, elective.faculty).border = thin_border
                key = elective.code + "||" + elective.title
                classroom = elective_assignment.get(key, "") or ""
                ws.cell(start_row + i, 6, classroom).border = thin_border
                ws.cell(start_row + i, 7, "").fill = PatternFill(
                    start_color=color_map.get("Elective", "FFFFFF"),
                    end_color=color_map.get("Elective", "FFFFFF"),
                    fill_type="solid",
                )
                ws.cell(start_row + i, 7).border = thin_border
                i += 1

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for col in ws.columns:
                max_length = 0
                try:
                    column = col[0].column_letter
                except Exception:
                    continue
                for cell in col:
                    try:
                        if cell.value is not None:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                ws.column_dimensions[column].width = max_length + 2

        wb.save(filename)
        print(f"Formatted student timetable saved in {filename}")

    def _generate_faculty_workbook(self, faculty_filename):
       
        faculty_set = set()
        for c in self.courses:
            raw = c.faculty
            if raw:
                for p in [x.strip() for x in raw.split("/") if x.strip()]:
                    faculty_set.add(p)
        for ent in self.scheduled_entries:
            fstr = ent.get("faculty")
            if fstr:
                for p in [x.strip() for x in fstr.split("/") if x.strip()]:
                    faculty_set.add(p)

       
        faculty_sheets = {}
        for f in faculty_set:
            faculty_sheets[f] = pd.DataFrame("    ", index=self.days, columns=self.slots)

        
        for ent in self.scheduled_entries:
            day = ent["day"]
            slot = ent["slot"]
            display = ent["display"]  
           
            if ent.get("faculty"):
                faculties = [p.strip() for p in ent["faculty"].split("/") if p.strip()]
            else:
                matched = [c for c in self.courses if c.code == ent.get("code")]
                faculties = []
                for m in matched:
                    faculties.extend([p.strip() for p in m.faculty.split("/") if p.strip()])

            for f in faculties:
                if f not in faculty_sheets:
                    faculty_sheets[f] = pd.DataFrame("FREE", index=self.days, columns=self.slots)
                faculty_sheets[f].at[day, slot] = display

        
        with pd.ExcelWriter(faculty_filename, engine="openpyxl") as writer:
            for f in sorted(faculty_sheets.keys()):
                safe_name = f[:31]
                faculty_sheets[f].to_excel(writer, sheet_name=safe_name, index=True)

        
        wb = load_workbook(faculty_filename)
        thin = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

        
        palette = ["FFC7CE", "C6EFCE", "FFEB9C", "BDD7EE", "D9EAD3", "F4CCCC", "D9D2E9", "FCE5CD", "C9DAF8", "EAD1DC"]
        color_map = {}
        color_index = 0

        for sheet in wb.sheetnames:
            ws = wb[sheet]
         
            for row in range(2, ws.max_row + 1):
                start_col = 2
                while start_col <= ws.max_column:
                    cell = ws.cell(row=row, column=start_col)
                    if cell.value and str(cell.value).strip() not in ["FREE", ""]:
                        raw_code = str(cell.value).split(" ")[0].rstrip("T")
                        if raw_code not in color_map:
                            color_map[raw_code] = palette[color_index % len(palette)]
                            color_index += 1
                        fill = PatternFill(start_color=color_map[raw_code], end_color=color_map[raw_code], fill_type="solid")
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
                            ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=start_col + merge_count)
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        for col_idx in range(start_col, start_col + merge_count + 1):
                            ws.cell(row=row, column=col_idx).border = thin
                        start_col += merge_count + 1
                    else:
                        cell.border = thin
                        start_col += 1

            
            for col in ws.columns:
                max_length = 0
                try:
                    column = col[0].column_letter
                except Exception:
                    continue
                for cell in col:
                    try:
                        if cell.value is not None:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                ws.column_dimensions[column].width = max_length + 2

        wb.save(faculty_filename)
        print(f"Saved faculty timetables to {faculty_filename} (with colors and merged cells)")

    def run_all_outputs(self, dept_name_prefix="CSE", student_filename=None, faculty_filename="faculty_timetable.xlsx"):
        
        if not student_filename:
            student_filename = f"{dept_name_prefix}_timetable.xlsx"

        
        self.scheduled_entries = []
        self.electives_by_sheet = {}
        self.elective_room_assignment = {}

        with pd.ExcelWriter(student_filename, engine="openpyxl") as writer:
           
            self.generate_timetable([c for c in self.courses if c.semester_half in ["1", "0"]], writer, "First_Half")
            self.generate_timetable([c for c in self.courses if c.semester_half in ["2", "0"]], writer, "Second_Half")

        
        wb = load_workbook(student_filename)
        for default in ["Sheet", "Sheet1"]:
            if default in wb.sheetnames and len(wb.sheetnames) > 1:
                wb.remove(wb[default])
        wb.save(student_filename)

        
        for sheet in wb.sheetnames:
            self._compute_elective_room_assignments_legally(sheet)

        
        self.format_student_timetable_with_legend(student_filename)

        
        self._generate_faculty_workbook(faculty_filename)


if __name__ == "__main__":
   
    departments = {
        "CSE": "data/CSE_courses.csv",
        "DSAI": "data/DSAI_courses.csv",
    }
    rooms_file = "data/rooms.csv"
    slots_file = "data/timeslots.csv"

    
    global_room_usage = {}

    
    combined_faculty_filename = "faculty_timetable.xlsx"
    
    all_scheduled_entries = []

    for dept_name, course_file in departments.items():
        print(f"\nGenerating student timetable for {dept_name}...")
        scheduler = Scheduler(slots_file, course_file, rooms_file, global_room_usage)
        student_file = f"{dept_name}_timetable.xlsx"
        scheduler.run_all_outputs(dept_name_prefix=dept_name, student_filename=student_file, faculty_filename=combined_faculty_filename)
        
        all_scheduled_entries.extend(scheduler.scheduled_entries)
        
        for k, v in scheduler.course_room_map.items():
            
            global_room_usage.setdefault("MAPPING", {})[k] = v

    
    combined_courses = []
    for dept_name, course_file in departments.items():
        df = pd.read_csv(course_file)
        for _, row in df.iterrows():
            combined_courses.append(Course(row))
    helper = Scheduler(slots_file, departments[list(departments.keys())[0]], rooms_file, global_room_usage)
    helper.courses = combined_courses
    helper.scheduled_entries = all_scheduled_entries 
    helper._generate_faculty_workbook(combined_faculty_filename)
    print("\nAll done. Student timetables and combined faculty timetable generated.")
