import unittest
from timetable_automation.main import Course

class TestCourse(unittest.TestCase):
    def test_course_parsing(self):
        row = {
            "Course_Code": "CS101",
            "Course_Title": "Intro to Programming",
            "Faculty": "Prof. X",
            "L-T-P-S-C": "3-1-0-0-4",
            "Semester_Half": "1",
            "Elective": "0"
        }
        course = Course(row)
        self.assertEqual(course.code, "CS101")
        self.assertEqual(course.title, "Intro to Programming")
        self.assertEqual((course.L, course.T, course.P), (3, 1, 0))
        self.assertFalse(course.is_elective)

    def test_invalid_format_defaults(self):
        row = {"Course_Code": "CS102", "L-T-P-S-C": "bad-data"}
        course = Course(row)
        self.assertEqual((course.L, course.T, course.P), (0, 0, 0))
