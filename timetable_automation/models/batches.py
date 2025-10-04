class Batch:
    def __init__(self, department, semester, total_students, max_batch_size):
        self.department = department         # e.g. "CSE"
        self.semester = semester             # e.g. 3
        self.total_students = total_students # e.g. 170
        self.max_batch_size = max_batch_size # e.g. 85

    def __repr__(self):
        return (f"Batch(Department={self.department}, Semester={self.semester}, "
                f"Total_Students={self.total_students}, MaxBatchSize={self.max_batch_size})")
