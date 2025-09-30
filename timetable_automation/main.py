from timetable_automation.services.timetable_generator import generate_timetable

def main():
    print("Welcome to Timetable Automation System")
    timetable = generate_timetable()
    print("Generated Timetable:", timetable)

if __name__ == "__main__":
    main()
