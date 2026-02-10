import unittest
from pathlib import Path

import pandas as pd

from timetable_automation.main import Scheduler


class TestElectiveRoomAlignment(unittest.TestCase):
    def setUp(self):
        self.test_data_dir = Path("tests/test_data")
        self.test_data_dir.mkdir(parents=True, exist_ok=True)

        self.slots_file = self.test_data_dir / "align_slots.csv"
        self.rooms_file = self.test_data_dir / "align_rooms.csv"
        self.courses_cse_file = self.test_data_dir / "align_courses_cse5.csv"
        self.courses_ece_file = self.test_data_dir / "align_courses_ece5.csv"

        pd.DataFrame(
            [
                {"Start_Time": "09:00", "End_Time": "10:00"},
            ]
        ).to_csv(self.slots_file, index=False)

        pd.DataFrame(
            [
                {"Room_ID": "C101", "Capacity": 96, "Type": "Classroom"},
                {"Room_ID": "C102", "Capacity": 96, "Type": "Classroom"},
            ]
        ).to_csv(self.rooms_file, index=False)

        pd.DataFrame(
            [
                {
                    "Course_Code": "EC364",
                    "Course_Title": "Semiconductor Device Modeling",
                    "L-T-P-S-C": "3-0-0-0-3",
                    "Faculty": "Dr. Pankaj Kumar",
                    "Semester_Half": "0",
                    "Elective": "1",
                    "Students": 71,
                    "basket": 1,
                }
            ]
        ).to_csv(self.courses_cse_file, index=False)

        pd.DataFrame(
            [
                {
                    "Course_Code": "EC364",
                    "Course_Title": "Semiconductor Device Modeling",
                    "L-T-P-S-C": "3-0-0-0-3",
                    "Faculty": "Dr. Pankaj Kumar",
                    "Semester_Half": "0",
                    "Elective": "1",
                    "Students": 71,
                    "basket": 2,
                }
            ]
        ).to_csv(self.courses_ece_file, index=False)

    def tearDown(self):
        for f in (
            self.slots_file,
            self.rooms_file,
            self.courses_cse_file,
            self.courses_ece_file,
        ):
            try:
                Path(f).unlink()
            except FileNotFoundError:
                pass

    def test_same_code_elective_uses_same_room_across_branches(self):
        shared_room_usage = {}
        shared_elective_slots = {}
        shared_elective_slot_usage = {}
        shared_elective_room_templates = {}
        shared_elective_room_usage = {}
        shared_elective_representatives = {}

        cse = Scheduler(
            str(self.slots_file),
            str(self.courses_cse_file),
            str(self.rooms_file),
            shared_room_usage,
            shared_elective_slots,
            dept_name="CSE-5-A",
            global_elective_slot_usage=shared_elective_slot_usage,
            global_elective_room_templates=shared_elective_room_templates,
            global_elective_room_usage=shared_elective_room_usage,
            global_elective_representatives=shared_elective_representatives,
        )

        ece = Scheduler(
            str(self.slots_file),
            str(self.courses_ece_file),
            str(self.rooms_file),
            shared_room_usage,
            shared_elective_slots,
            dept_name="ECE-5",
            global_elective_slot_usage=shared_elective_slot_usage,
            global_elective_room_templates=shared_elective_room_templates,
            global_elective_room_usage=shared_elective_room_usage,
            global_elective_representatives=shared_elective_representatives,
        )

        sheet_name = "First_Half"
        slot = cse.slots[0]

        cse.electives_by_sheet[sheet_name] = [(1, cse.courses[0])]
        cse.scheduled_entries = [
            {
                "sheet": sheet_name,
                "day": "Monday",
                "slot": slot,
                "code": "Elective_1",
                "display": "Elective_1",
                "faculty": cse.courses[0].faculty,
                "room": "",
            }
        ]
        cse._compute_elective_room_assignments_legally(sheet_name)

        ece.electives_by_sheet[sheet_name] = [(2, ece.courses[0])]
        ece.scheduled_entries = [
            {
                "sheet": sheet_name,
                "day": "Monday",
                "slot": slot,
                "code": "Elective_2",
                "display": "Elective_2",
                "faculty": ece.courses[0].faculty,
                "room": "",
            }
        ]
        ece._compute_elective_room_assignments_legally(sheet_name)

        cse_room = cse.elective_room_assignment[sheet_name]["Elective_1||Semiconductor Device Modeling"]
        ece_room = ece.elective_room_assignment[sheet_name]["Elective_2||Semiconductor Device Modeling"]
        self.assertEqual(cse_room, ece_room)

        code_template_key = ("5", sheet_name, "__CODE__", "EC364")
        self.assertIn(code_template_key, shared_elective_room_templates)


if __name__ == "__main__":
    unittest.main()
