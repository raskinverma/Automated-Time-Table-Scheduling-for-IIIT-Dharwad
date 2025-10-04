class Course:
    def __init__(self, department, semester, course_code, course_name, L, T, P, S, C, faculty):
        self.department = department   
        self.semester = semester        
        self.course_code = course_code  
        self.course_name = course_name  
        self.L = int(L)              
        self.T = int(T)               
        self.P = int(P)            
        self.S = int(S)               
        self.C = int(C)                 
        self.faculty = faculty         

    def __repr__(self):
        return (f"Course({self.department}, Sem {self.semester}, {self.course_code}, "
                f"{self.course_name}, L={self.L}, T={self.T}, P={self.P}, S={self.S}, C={self.C}, "
                f"Faculty={self.faculty})")
