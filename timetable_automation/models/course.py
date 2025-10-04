class Course:
    def __init__(self, department, semester, course_code, course_name, ltp_sc, faculty):
        self.department = department      
        self.semester = semester            
        self.course_code = course_code    
        self.course_name = course_name     
        self.ltp_sc = ltp_sc             
        self.faculty = faculty            

    def __repr__(self):
        return (f"Course({self.department}, Sem {self.semester}, {self.course_code}, "
                f"{self.course_name}, {self.ltp_sc}, {self.faculty})")
