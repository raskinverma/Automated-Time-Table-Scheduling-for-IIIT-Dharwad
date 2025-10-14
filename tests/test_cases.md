# TEST CASES — Timetable Automation Project

---

## 1. `Course` Class

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|
| Valid row with all fields | `Course(row)` | Parses course details correctly | `code="CS101"`, `title="Intro to Programming"`, `(L,T,P)=(3,1,0)`, `is_elective=False` |
| Invalid L-T-P-S-C format | `Course(row)` | Handles malformed L-T-P-S-C string gracefully | `(L,T,P)=(0,0,0)` |
| Missing optional fields | `Course(row)` | Uses defaults when Faculty or Semester_Half is missing | No crash; default empty/zero values |
| Trailing spaces | `Course(row)` | Trims leading/trailing whitespace | Cleaned string values |
| Elective as string `"1"` | `Course(row)` | Converts "1"/"0" to boolean correctly | `is_elective=True` |

---

## 2. `Scheduler` — Basic Functions

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|
| `"09:00-10:00"` | `_slot_duration(slot)` | Calculates 1-hour slot duration | `1.0` |
| `"10:00-11:30"` | `_slot_duration(slot)` | Calculates 1.5-hour slot duration | `1.5` |
| `df` with `"09:00-10:00"` used | `_get_free_blocks(df, "Monday")` | Finds available continuous free blocks | Returns list of free slot groups |
| All slots empty | `_get_free_blocks(df, "Monday")` | Handles full-day free scenario | Returns one large free block |
| Invalid/overlapping times | `_slot_duration(slot)` | Handles bad inputs gracefully | Returns `0.0` or safe error handling |

---

## 3. `Scheduler` — Session Allocation

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|
| Free timetable, 1-hour lecture | `_allocate_session()` | Allocates one session successfully | Returns `True` |
| Slot already occupied | `_allocate_session()` | Avoids double-booking | Returns `False` |
| Room conflict | `_allocate_session()` | Prevents assigning same room twice in a slot | Returns `False` |
| Elective session | `_allocate_session()` | Skips room requirement | Returns `True` with empty room |
| Lab allocation | `_allocate_session()` | Ensures only one lab per day | Respects `daily_lab_done` constraint |

---

## 4. `Scheduler` — Elective Room Assignment

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|
| One elective course | `_compute_elective_room_assignments_legally("Sheet1")` | Assigns room to elective | Adds `"Sheet1"` in `elective_room_assignment` |
| No electives present | `_compute_elective_room_assignments_legally("Sheet1")` | Skips processing cleanly | Returns without error |
| Multiple electives, same slot | `_compute_elective_room_assignments_legally("Sheet1")` | Ensures unique rooms | Distinct room assignment per elective |
| All rooms occupied | `_compute_elective_room_assignments_legally("Sheet1")` | Falls back to available or ordered rooms | Graceful assignment with no duplicates |

---

## 5. `Scheduler` — End-to-End Generation

| Test Case Input | Function | Description | Expected Output |
|-----------------|-----------|--------------|-----------------|
| Small CSVs for slots, courses, rooms | `generate_timetable()` | Generates valid timetable for all inputs | Excel output generated |
| Duplicate room test | `generate_timetable()` | Ensures no duplicate room usage | No duplicate room entries per slot |
| Elective included | `generate_timetable()` | Adds placeholder “Elective” course | Timetable includes elective rows |
| Empty course data | `generate_timetable()` | Handles no-course scenario gracefully | Skips sheet creation |
| Writer closed properly | `generate_timetable()` | Closes ExcelWriter cleanly | `.xlsx` file exists and valid |

---
