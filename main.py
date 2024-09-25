import pandas as pd
from datetime import time, datetime, timedelta

# Load the Excel file
file_path = 'BB_IOT.xlsx'  # Ganti dengan path file Excel Anda
df = pd.read_excel(file_path)

# Pengaturan
START_TIME = time(8, 0)  # Waktu mulai kelas
END_TIME = time(16, 50)  # Waktu akhir kelas
MAX_CONFLICTS = 5  # Maksimum konflik
DURATION = timedelta(hours=2, minutes=50)  # Durasi kelas
INTERVAL_MINUTES = 20  # Interval pencarian dalam menit

# Menyiapkan slot waktu
def generate_time_slots(start, end, interval):
    slots = []
    current_time = datetime.combine(datetime.today(), start)
    end_time = datetime.combine(datetime.today(), end)
    
    while current_time <= end_time:
        slots.append(current_time.time())
        current_time += timedelta(minutes=interval)
    
    return slots

# Membuat daftar slot waktu
time_slots = generate_time_slots(START_TIME, END_TIME, INTERVAL_MINUTES)

# Inisialisasi dictionary untuk mencatat ketersediaan
availability = {
    'SENIN': {slot: {'count': 0, 'students': []} for slot in time_slots},
    'SELASA': {slot: {'count': 0, 'students': []} for slot in time_slots},
    'RABU': {slot: {'count': 0, 'students': []} for slot in time_slots},
    'KAMIS': {slot: {'count': 0, 'students': []} for slot in time_slots},
    'JUMAT': {slot: {'count': 0, 'students': []} for slot in time_slots},
    'SABTU': {slot: {'count': 0, 'students': []} for slot in time_slots},
}

# Fungsi untuk menandai slot waktu yang sibuk
def mark_busy_slots(student_name, day, start_str, end_str):
    start_time_class = pd.to_datetime(start_str, format='%H:%M').time()
    end_time_class = pd.to_datetime(end_str, format='%H:%M').time()
    
    for slot in time_slots:
        # Jika slot berada dalam rentang waktu kelas, tandai sebagai sibuk
        if start_time_class <= slot < end_time_class:
            availability[day][slot]['count'] += 1
            availability[day][slot]['students'].append(student_name)

# Proses data jadwal untuk semua mahasiswa
for index, row in df.iterrows():
    mark_busy_slots(row['NIM'], row['Disp_Hari'], row['Disp_Jam'].split(' - ')[0], row['Disp_Jam'].split(' - ')[1])

# Fungsi untuk mencari slot kosong
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
                
                # Cek jika kita sudah mendapatkan cukup slot untuk 2 jam 50 menit
                if len(temp_free) * (INTERVAL_MINUTES / 60) >= 2.83:  # 2 jam 50 menit dalam jam
                    conflicting_students = []
                    for t in temp_free:
                        conflicting_students += slots[t]['students']
                    conflicting_students = list(set(conflicting_students))  # Hapus duplikat
                    
                    # Batasi jumlah mahasiswa konflik hingga MAX_CONFLICTS
                    conflicting_students = conflicting_students[:MAX_CONFLICTS]
                    
                    # Tambahkan informasi slot kosong
                    free_slots[day].append((temp_free[0], temp_free[-1], conflicting_students))
                    
                    # Lewati sisa slot dalam periode ini
                    i += len(temp_free) - 1
                    temp_free = []
            else:
                temp_free = []

            i += 1  # Lanjut ke slot berikutnya

    return free_slots

# Dapatkan slot kosong untuk semua mahasiswa
common_free_slots = find_common_free_slots()

# Tampilkan slot kosong dan mahasiswa yang konflik
for day, slots in common_free_slots.items():
    print(f"Slot kosong pada {day}:")
    for start_slot, end_slot, students_conflict in slots:
        end_time_slot = (datetime.combine(datetime.today(), start_slot) + DURATION).time()  # Hitung waktu akhir
        print(f"  {start_slot.strftime('%H:%M')} - {end_time_slot.strftime('%H:%M')}")
        if students_conflict:
            print(f"    Mahasiswa konflik (maksimal {MAX_CONFLICTS}):")
            for student in students_conflict:
                print(f"      - {student}")  # Cetak setiap mahasiswa konflik di baris baru
