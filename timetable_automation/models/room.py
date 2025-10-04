class Room:
    def __init__(self, room_id, capacity, rtype, facilities):
        self.room_id = room_id       
        self.capacity = capacity    
        self.rtype = rtype          
        self.facilities = facilities 

    def __repr__(self):
        return (f"Room({self.room_id}, capacity={self.capacity}, "
                f"type={self.rtype}, facilities={self.facilities})")
