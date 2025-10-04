class Faculty:
    def __init__(self, faculty_id, name):
        self.faculty_id = faculty_id  # e.g. "F001"
        self.name = name              # e.g. "Dr. Abdul Wahid"

    def __repr__(self):
        return f"Faculty({self.faculty_id}, {self.name})"
