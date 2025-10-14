import unittest
from timetable_automation.main import Scheduler, Course

class TestElectiveAssignment(unittest.TestCase):
    def setUp(self):
        self.sched = Scheduler("tests/test_data/slots.csv", "tests/test_data/courses.csv", "tests/test_data/rooms.csv", {})

    def test_elective_room_assignment(self):
        self.sched.electives_by_sheet["Sheet1"] = [Course({
            "Course_Code": "EL101",
            "Course_Title": "Elective Course",
            "Faculty": "Prof B",
            "L-T-P-S-C": "1-0-0-0-1",
            "Semester_Half": "1",
            "Elective": 1
        })]
        self.sched.scheduled_entries = [
            {"sheet": "Sheet1", "day": "Monday", "slot": "09:00-10:00", "code": "Elective"}
        ]
        self.sched._compute_elective_room_assignments_legally("Sheet1")
        self.assertIn("Sheet1", self.sched.elective_room_assignment)
