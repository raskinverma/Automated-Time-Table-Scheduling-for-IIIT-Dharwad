import unittest
import pandas as pd
from timetable_automation.main import Scheduler

class TestEndToEnd(unittest.TestCase):
    def test_end_to_end_schedule(self):
        # Create small dummy data
        slots = pd.DataFrame([
            {"Start_Time": "09:00", "End_Time": "10:00"},
            {"Start_Time": "10:00", "End_Time": "11:00"}
        ])
        courses = pd.DataFrame([
            {"Course_Code": "C1", "Course_Title": "X", "Faculty": "F1", "L-T-P-S-C": "1-0-0-0-1", "Semester_Half": "1", "Elective": "0"},
            {"Course_Code": "C2", "Course_Title": "Y", "Faculty": "F2", "L-T-P-S-C": "1-0-0-0-1", "Semester_Half": "1", "Elective": "0"}
        ])
        rooms = pd.DataFrame([
            {"Room_ID": "R1", "Type": "classroom"},
            {"Room_ID": "R2", "Type": "classroom"}
        ])

        slots.to_csv("tests/test_data/e2e_slots.csv", index=False)
        courses.to_csv("tests/test_data/e2e_courses.csv", index=False)
        rooms.to_csv("tests/test_data/e2e_rooms.csv", index=False)

        sched = Scheduler("tests/test_data/e2e_slots.csv", "tests/test_data/e2e_courses.csv", "tests/test_data/e2e_rooms.csv", {})
        writer = pd.ExcelWriter("tests/test_data/test_output.xlsx", engine="openpyxl")
        sched.generate_timetable(sched.courses, writer, "TestSheet")
        writer.close()

        # Validate: no duplicate room usage
        for day, slot_map in sched.global_room_usage.items():
            for slot, rooms_used in slot_map.items():
                self.assertEqual(len(set(rooms_used)), len(rooms_used))
