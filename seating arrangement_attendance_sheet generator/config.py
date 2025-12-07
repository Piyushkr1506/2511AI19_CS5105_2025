# Put room inventory and capacities here or load from a file
ROOMS = [
# (building, room_code, capacity)
("B1", "6101", 50),
("B1", "6102", 45),
("B1", "6103", 40),
("B2", "10502", 60),
("B2", "10503", 55),
("Auditorium", "AUD1", 200),
]


# Adjacency groups — to prioritize close rooms
ADJACENCY = {
'B1': ['6101', '6102', '6103'],
'B2': ['10502', '10503'],
'Auditorium': ['AUD1'],
}