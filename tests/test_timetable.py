import unittest
from timetable_automation.services.timetable_generator import generate_timetable

class TestTimetable(unittest.TestCase):
    def test_generate(self):
        timetable = generate_timetable()
        self.assertIn("Monday", timetable)

if __name__ == "__main__":
    unittest.main()
