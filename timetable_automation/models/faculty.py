class Faculty:
    def __init__(self, faculty_id, name):
        self.faculty_id = faculty_id 
        self.name = name              

    def __repr__(self):
        return f"Faculty({self.faculty_id}, {self.name})"
