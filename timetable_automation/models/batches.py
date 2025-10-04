class Batch:
    def __init__(self, department, semester, total_students, max_batch_size):
        self.department = department      
        self.semester = semester        
        self.total_students = total_students 
        self.max_batch_size = max_batch_size 

    def __repr__(self):
        return (f"Batch(Department={self.department}, Semester={self.semester}, "
                f"Total_Students={self.total_students}, MaxBatchSize={self.max_batch_size})")
