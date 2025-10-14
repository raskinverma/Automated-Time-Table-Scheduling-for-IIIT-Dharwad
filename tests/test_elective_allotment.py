import unittest
import pandas as pd
from io import StringIO
from timetable_automation.main import Scheduler, Course 

class TestElectiveAllotment(unittest.TestCase):
    def setUp(self):
        # Mock timeslots CSV
        slots_csv = StringIO(
            "Start_Time,End_Time\n"
            "08:00,09:30\n"
            "09:30,11:00\n"
            "11:00,12:30\n"
        )
        self.slots_file = "tests/test_data/mock_slots.csv"
        pd.read_csv(slots_csv).to_csv(self.slots_file, index=False)

        # Mock rooms CSV
        rooms_csv = StringIO(
            "Room_ID,Type\n"
            "R101,Classroom\n"
            "R102,Lab\n"
        )
        self.rooms_file = "tests/test_data/mock_rooms.csv"
        pd.read_csv(rooms_csv).to_csv(self.rooms_file, index=False)

        # Mock courses CSV
        courses_csv = StringIO(
            "Course_Code,Course_Title,L-T-P-S-C,Faculty,Semester_Half,Elective\n"
            "CS101,Intro,3-0-0-0-3,Alice,1,1\n"
            "CS102,DS,3-0-0-0-3,Bob,1,0\n"
        )
        self.courses_file = "tests/test_data/mock_courses.csv"
        pd.read_csv(courses_csv).to_csv(self.courses_file, index=False)

        # Global room usage
        self.global_room_usage = {}

    def test_elective_allocation(self):
        scheduler = Scheduler(self.slots_file, self.courses_file, self.rooms_file, self.global_room_usage)
        
        # Run timetable generation (without writing to Excel)
        writer = pd.ExcelWriter("tests/test_data/mock_output.xlsx", engine="openpyxl")
        scheduler.generate_timetable(scheduler.courses, writer, "TestSheet")
        writer.close()

        # Electives should be scheduled
        electives = [c for c in scheduler.courses if c.is_elective]
        self.assertTrue(len(electives) > 0)

        elective_entries = [e for e in scheduler.scheduled_entries if e["code"] == "Elective"]
        self.assertTrue(len(elective_entries) > 0, "Electives should be scheduled in the timetable")

        # Core courses should not appear as 'Elective'
        core_entries = [e for e in scheduler.scheduled_entries if e["code"] != "Elective"]
        self.assertTrue(all(not e["code"].startswith("Elective") for e in core_entries))

        # Check faculty assignment exists for electives
        for e in elective_entries:
            self.assertIn(e["faculty"], [c.faculty for c in scheduler.courses if c.is_elective or c.code == "Elective"])

if __name__ == "__main__":
    unittest.main()
