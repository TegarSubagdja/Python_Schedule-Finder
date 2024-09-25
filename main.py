import pandas as pd
from datetime import time, datetime, timedelta

# Load the Excel file
file_path = 'AA_IOT.xlsx'  # Change this to the path of your Excel file
df = pd.read_excel(file_path)

# Define the time range to check availability (7 AM to 2 PM)
start_time = pd.to_datetime('07:00', format='%H:%M').time()
end_time = pd.to_datetime('14:00', format='%H:%M').time()

# Time slots in 30-minute intervals
time_slots = [time(h, m) for h in range(7, 15) for m in (0, 30)]

# Initialize the dictionary to store busy status and student names for each time slot for each day
availability = {
    'SENIN': {slot: {'count': 0, 'students': []} for slot in time_slots},
    'SELASA': {slot: {'count': 0, 'students': []} for slot in time_slots},
    'RABU': {slot: {'count': 0, 'students': []} for slot in time_slots},
    'KAMIS': {slot: {'count': 0, 'students': []} for slot in time_slots},
    'JUMAT': {slot: {'count': 0, 'students': []} for slot in time_slots},
    'SABTU': {slot: {'count': 0, 'students': []} for slot in time_slots},
}

# Maximum allowed conflict
MAX_CONFLICTS = 5

# Function to mark time slots as busy and track student names
def mark_busy_slots(student_name, day, start_str, end_str):
    start_time_class = pd.to_datetime(start_str, format='%H:%M').time()
    end_time_class = pd.to_datetime(end_str, format='%H:%M').time()
    
    for slot in time_slots:
        # If the slot falls within the class time, mark it as busy and add the student name
        if start_time_class <= slot < end_time_class:
            availability[day][slot]['count'] += 1
            availability[day][slot]['students'].append(student_name)

# Process the schedule data for all students
for index, row in df.iterrows():
    mark_busy_slots(row['NIM'], row['Disp_Hari'], row['Disp_Jam'].split(' - ')[0], row['Disp_Jam'].split(' - ')[1])

# Function to add 30 minutes to a given time slot
def add_30_minutes(slot):
    slot_time = datetime.combine(datetime.today(), slot)
    slot_time_plus_30 = slot_time + timedelta(minutes=30)
    return slot_time_plus_30.time()

# Function to find free time slots (at least 2 hours 50 minutes, up to 5 conflicts allowed)
def find_common_free_slots():
    free_slots = {}

    for day, slots in availability.items():
        free_slots[day] = []
        temp_free = []
        i = 0
        
        while i < len(time_slots):
            slot = time_slots[i]
            slot_data = slots[slot]

            if slot_data['count'] <= MAX_CONFLICTS:
                temp_free.append(slot)
                # Check if we have at least 5 consecutive free slots (i.e. 2 hours 50 minutes)
                if len(temp_free) >= 5:
                    # Add the start and end time for the free block along with conflicting students' names
                    conflicting_students = []
                    for t in temp_free:
                        conflicting_students += slots[t]['students']
                    conflicting_students = list(set(conflicting_students))  # Remove duplicates
                    free_slots[day].append((temp_free[0], temp_free[-1], conflicting_students))  # Start, end, and conflicts
                    
                    # Skip over the remaining slots in this 2-hour 50-minute period
                    i += len(temp_free) - 1  # Move past this block
                    temp_free = []  # Reset temp_free after collecting a block
            else:
                temp_free = []  # Reset when the slot is busy

            i += 1  # Move to the next slot

    return free_slots

# Get the common free slots for all students
common_free_slots = find_common_free_slots()

# Display the free time slots and conflicting students below each free slot group
for day, slots in common_free_slots.items():
    print(f"Free slots on {day}:")
    for start_slot, end_slot, students in slots:
        end_time_slot = add_30_minutes(end_slot)  # Calculate the end time
        print(f"  {start_slot.strftime('%H:%M')} - {end_time_slot.strftime('%H:%M')}")
        if students:
            print(f"    Conflicts with:")
            for student in students:
                print(f"      - {student}")  # Print each conflicting student on a new line
