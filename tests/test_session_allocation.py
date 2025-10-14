import unittest
import pandas as pd
from timetable_automation.main import Scheduler

class TestSessionAllocation(unittest.TestCase):
    def setUp(self):
        pd.DataFrame([
            {"Start_Time": "09:00", "End_Time": "10:00"},
            {"Start_Time": "10:00", "End_Time": "11:00"}
        ]).to_csv("tests/test_data/slots.csv", index=False)

        pd.DataFrame([
            {"Course_Code": "CS101", "Course_Title": "Intro", "Faculty": "Prof X", "L-T-P-S-C": "1-0-0-0-1", "Semester_Half": "1", "Elective": "0"},
        ]).to_csv("tests/test_data/courses.csv", index=False)

        pd.DataFrame([
            {"Room_ID": "R1", "Type": "classroom"},
            {"Room_ID": "R2", "Type": "lab"}
        ]).to_csv("tests/test_data/rooms.csv", index=False)

        self.sched = Scheduler("tests/test_data/slots.csv", "tests/test_data/courses.csv", "tests/test_data/rooms.csv", {})

    def test_allocate_basic_session(self):
        df = pd.DataFrame("", index=self.sched.days, columns=self.sched.slots)
        success = self.sched._allocate_session(
            df,
            {d: [] for d in self.sched.days},
            {d: False for d in self.sched.days},
            "Monday",
            "Prof X",
            "CS101",
            1.0,
            "L",
            False,
            "TestSheet"
        )
        self.assertTrue(success)
