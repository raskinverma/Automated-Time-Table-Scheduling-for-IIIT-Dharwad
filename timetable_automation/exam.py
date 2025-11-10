import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side, PatternFill
from datetime import datetime, timedelta

SLOT_LABELS = ["09:00-12:00", "14:00-17:00"]
MAX_GLOBAL_EXAMS_PER_DAY = 3
MAX_EXAMS_PER_GROUP_PER_DAY = 1
DEFAULT_START_DATE = "2025-11-20"

def invigilators_needed(capacity):
    return 2 if capacity >= 200 else 1

def extract_semester_id(group_name: str) -> str:
    m = re.search(r"(\d+)", str(group_name))
    return m.group(1) if m else str(group_name)

class Course:
    def __init__(self, row, group_name):
        self.group = group_name
        self.code = str(row["Course_Code"]).strip()
        self.title = str(row.get("Course_Title", self.code)).strip()
        try:
            self.students = int(str(row.get("Students", "0")).strip())
        except:
            self.students = 0
        flag = str(row.get("Elective", "0")).strip()
        self.is_elective = flag in ("1", "true", "True", "YES", "yes")

class ExamScheduler:
    def __init__(self, rooms_file, departments, faculty_file, start_date=DEFAULT_START_DATE):
        self.rooms_df = pd.read_csv(rooms_file)
        self.departments = departments
        self.invig_df = pd.read_csv(faculty_file)
        self.start_date = datetime.strptime(start_date, "%Y-%m-%d").date()
        self.groups = list(departments.keys())
        self.invigilators = sorted([str(n).strip() for n in self.invig_df["Name"] if str(n).strip()])
        self.rooms = self._load_rooms()
        self.courses = self._load_courses()
        self.room_remaining = {}
        self.group_daily = {}
        self.global_daily = {}
        self.used_rooms = {}
        self.scheduled = []
        self.unscheduled = []
        self.invig_assignments = []
        self._inv_idx = 0

    def _load_rooms(self):
        rooms = []
        for _, r in self.rooms_df.iterrows():
            rid = str(r["Room_ID"]).strip()
            t = str(r["Type"]).strip().lower()
            cap_raw = str(r["Capacity"]).strip()
            try:
                cap = int(cap_raw)
            except:
                cap = 0
            if cap <= 0:
                continue
            if "lab" in t or "library" in t:
                continue
            usable = cap // 2
            if usable <= 0:
                continue
            rooms.append({"Room_ID": rid, "Capacity": cap, "Usable": usable})
        rooms.sort(key=lambda x: (-x["Usable"], x["Room_ID"]))
        return rooms

    def _load_courses(self):
        out = {}
        for g, file in self.departments.items():
            df = pd.read_csv(file)
            lst = [Course(r, g) for _, r in df.iterrows() if int(str(r.get("Students", 0))) > 0]
            lst.sort(key=lambda c: (-c.students, c.code))
            out[g] = lst
        return out

    def _ensure_date(self, date):
        if date not in self.room_remaining:
            self.room_remaining[date] = {s: {r["Room_ID"]: r["Usable"] for r in self.rooms} for s in SLOT_LABELS}
        if date not in self.used_rooms:
            self.used_rooms[date] = {s: set() for s in SLOT_LABELS}
        if date not in self.group_daily:
            self.group_daily[date] = {g: 0 for g in self.groups}
        if date not in self.global_daily:
            self.global_daily[date] = 0

    def _alloc_rooms(self, date, slot, need):
        remaining = self.room_remaining[date][slot]
        room_order = sorted(self.rooms, key=lambda x: (-remaining.get(x["Room_ID"], 0), x["Room_ID"]))
        alloc = []
        total = 0
        for r in room_order:
            rid = r["Room_ID"]
            avail = remaining.get(rid, 0)
            if avail <= 0:
                continue
            take = min(avail, need - total)
            if take > 0:
                alloc.append((rid, take))
                total += take
            if total >= need:
                break
        return alloc if total >= need else None

    def _book_alloc(self, date, slot, alloc):
        for rid, cnt in alloc:
            self.room_remaining[date][slot][rid] -= cnt
            self.used_rooms[date][slot].add(rid)

    def _place_regular_course(self, course, date, slot):
        if self.global_daily[date] >= MAX_GLOBAL_EXAMS_PER_DAY:
            return False
        if self.group_daily[date][course.group] >= MAX_EXAMS_PER_GROUP_PER_DAY:
            return False
        alloc = self._alloc_rooms(date, slot, course.students)
        if alloc is None:
            return False
        self._book_alloc(date, slot, alloc)
        self.group_daily[date][course.group] += 1
        self.global_daily[date] += 1
        self.scheduled.append({
            "Date": date.strftime("%Y-%m-%d"),
            "Slot": slot,
            "Group": course.group,
            "Course_Code": course.code,
            "Course_Title": course.title,
            "Students": course.students,
            "Allocations": "; ".join([f"{rid}:{cnt}" for rid, cnt in alloc]),
            "Tag": "REG"
        })
        return True

    def _plan_electives_by_semester(self):
        sem_to_electives = {}
        sem_to_groups = {}
        for g in self.groups:
            sem = extract_semester_id(g)
            for c in self.courses[g]:
                if c.is_elective:
                    sem_to_electives.setdefault(sem, []).append(c)
                    sem_to_groups.setdefault(sem, set()).add(g)
        return sem_to_electives, sem_to_groups


    def _schedule_elective_block(self, sem, electives, groups_for_sem, start_day_offset, preferred_slot_index):
        day = start_day_offset

        while day < 300:
            date = self.start_date + timedelta(days=day)
            self._ensure_date(date)

            if self.global_daily[date] >= MAX_GLOBAL_EXAMS_PER_DAY:
                day += 1
                continue

            if any(self.group_daily[date][g] >= MAX_EXAMS_PER_GROUP_PER_DAY for g in groups_for_sem):
                day += 1
                continue

            slot = SLOT_LABELS[preferred_slot_index % len(SLOT_LABELS)]

            total_students = sum(c.students for c in electives)
            combined_alloc = self._alloc_rooms(date, slot, total_students)
            if combined_alloc is None:
                day += 1
                continue

            self._book_alloc(date, slot, combined_alloc)

            self.global_daily[date] += 1
            for g in groups_for_sem:
                self.group_daily[date][g] += 1

            ordered_rooms = [(rid, cnt) for rid, cnt in combined_alloc]

            for c in electives:
                need = c.students
                per_elec_alloc = []
                for rid, seats in ordered_rooms:
                    if need <= 0:
                        break
                    take = min(need, seats)
                    if take > 0:
                        per_elec_alloc.append((rid, take))
                        need -= take

                alloc_text = "; ".join(f"{rid}:{cnt}" for rid, cnt in per_elec_alloc)

                self.scheduled.append({
                    "Date": date.strftime("%Y-%m-%d"),
                    "Slot": slot,
                    "Group": c.group,
                    "Course_Code": c.code,
                    "Course_Title": c.title,
                    "Students": c.students,
                    "Allocations": alloc_text,
                    "Tag": f"ELECTIVE-S{sem}"
                })

            return day + 1

        return day

    def _remove_scheduled_electives_from_pool(self):
        for g in self.groups:
            self.courses[g] = [c for c in self.courses[g] if not c.is_elective]

    def _all_done(self):
        return all(len(self.courses[g]) == 0 for g in self.groups)

    def generate(self):
        sem_to_electives, sem_to_groups = self._plan_electives_by_semester()
        sem_list = sorted(sem_to_electives.keys(), key=lambda x: int(re.sub(r"\D", "", x)) if re.search(r"\d", x) else 0)

        day_cursor = 0
        for i, sem in enumerate(sem_list):
            electives = sem_to_electives[sem]
            if electives:
                groups_for_sem = sem_to_groups.get(sem, set())
                day_cursor = self._schedule_elective_block(sem, electives, groups_for_sem, day_cursor, i % 2)

        self._remove_scheduled_electives_from_pool()

        day = 0
        while not self._all_done() and day < 300:
            date = self.start_date + timedelta(days=day)
            self._ensure_date(date)

            pending = [g for g in self.groups if any(not c.is_elective for c in self.courses[g])]
            pending = pending[:MAX_GLOBAL_EXAMS_PER_DAY]

            slots_cycle = SLOT_LABELS * 2
            si = 0
            for g in pending:
                left = [c for c in self.courses[g] if not c.is_elective]
                if not left:
                    continue
                c = left[0]
                slot = slots_cycle[si]
                if self._place_regular_course(c, date, slot):
                    self.courses[g].remove(c)
                    si += 1

            day += 1

        for g in self.groups:
            for c in self.courses[g]:
                self.unscheduled.append({"Group": g, "Course_Code": c.code, "Course_Title": c.title, "Students": c.students})

        self._assign_invigilators()

    def _assign_invigilators(self):
        for d in sorted(self.used_rooms.keys()):
            assigned_today = set()
            for slot in SLOT_LABELS:
                rooms = sorted(list(self.used_rooms[d][slot]))
                for rid in rooms:
                    cap = next(r["Capacity"] for r in self.rooms if r["Room_ID"] == rid)
                    k = invigilators_needed(cap)

                    picks = []
                    while len(picks) < k:
                        name = self.invigilators[self._inv_idx % len(self.invigilators)]
                        self._inv_idx += 1
                        if name not in assigned_today:
                            picks.append(name)
                            assigned_today.add(name)
                        if len(assigned_today) == len(self.invigilators):
                            break

                    exam_names = []
                    date_str = d.strftime("%Y-%m-%d")
                    for rec in self.scheduled:
                        if rec["Date"] == date_str and rec["Slot"] == slot:
                            for a in rec["Allocations"].split(";"):
                                if rid == a.split(":")[0]:
                                    exam_names.append(f"{rec['Course_Code']} - {rec['Course_Title']}")
                                    break

                    self.invig_assignments.append({
                        "Date": date_str,
                        "Slot": slot,
                        "Room_ID": rid,
                        "Exam": " | ".join(sorted(set(exam_names))),
                        "Invigilators": ", ".join(picks)
                    })


    def _df_group(self, recs):
        dates = sorted({r["Date"] for r in recs})
        df = pd.DataFrame("", index=dates, columns=SLOT_LABELS)
        for r in recs:
            txt = f"{r['Course_Code']} - {r['Course_Title']} [{r['Students']}] @ {r['Allocations']}"
            cell = df.at[r["Date"], r["Slot"]]
            df.at[r["Date"], r["Slot"]] = (cell + "\n" if cell else "") + txt
        df.index.name = "Date"
        return df

    def _fmt(self, file):
        wb = load_workbook(file)
        thin = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        gray = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

        for s in wb.sheetnames:
            ws = wb[s]
            for row in ws.iter_rows():
                for cell in row:
                    cell.border = thin
                    if cell.row == 1 or cell.column == 1:
                        cell.fill = gray
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            for col in ws.columns:
                mx = 12
                for cell in col:
                    if cell.value:
                        mx = max(mx, len(str(cell.value)))
                ws.column_dimensions[col[0].column_letter].width = min(mx + 4, 70)

            ws.row_dimensions[1].height = 24

        wb.save(file)

    def export(self, out="exam_timetables.xlsx", uns="unscheduled_exams.xlsx", invig="invigilation.xlsx"):
        with pd.ExcelWriter(out, engine="openpyxl") as w:
            dfm = pd.DataFrame(self.scheduled)
            if dfm.empty:
                dfm = pd.DataFrame(columns=["Date","Slot","Group","Course_Code","Course_Title","Students","Allocations","Tag"])
            dfm.sort_values(by=["Date","Slot","Group","Course_Code"]).to_excel(w, sheet_name="Master", index=False)

            for g in self.groups:
                rec = [r for r in self.scheduled if r["Group"] == g]
                df = self._df_group(rec)
                if df.empty:
                    df = pd.DataFrame("", index=[self.start_date.strftime("%Y-%m-%d")], columns=SLOT_LABELS)
                    df.index.name = "Date"
                df.to_excel(w, sheet_name=g[:31], index=True)

        self._fmt(out)

        if self.unscheduled:
            pd.DataFrame(self.unscheduled).to_excel(uns, index=False)

        if self.invig_assignments:
            pd.DataFrame(self.invig_assignments).sort_values(by=["Date","Slot","Room_ID"]).to_excel(invig, index=False)

def run_example():
    departments = {
        "CSE-3": "data/exam_data/CSE_3.csv",
    }
    rooms = "data/exam_data/rooms.csv"
    faculty = "data/exam_data/faculty.csv"
    s = ExamScheduler(rooms, departments, faculty)
    s.generate()
    s.export()

if __name__ == "__main__":
    run_example()
