import cProfile
import pstats
from timetable_automation.main import Scheduler

def main():
    slots_file = "data/timeslots.csv"
    courses_file = "data/CSE_courses-A.csv"
    rooms_file = "data/rooms.csv"
    global_room_usage = {}

    scheduler = Scheduler(slots_file, courses_file, rooms_file, global_room_usage)
    scheduler.run_all_outputs(
        dept_name_prefix="CSE-A",
        student_filename="CSE-A_timetable_profiled.xlsx",
        faculty_filename="faculty_timetable_profiled.xlsx"
    )

if __name__ == "__main__":
    with cProfile.Profile() as prof:
        main()

    stats = pstats.Stats(prof)
    stats.sort_stats(pstats.SortKey.TIME)
    stats.print_stats(100)
    stats.dump_stats("timetable.prof")  # <-- Add this line
