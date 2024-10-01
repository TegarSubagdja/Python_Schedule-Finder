import requests
import time
import pandas as pd

# URL dasar yang akan diakses
base_url = "https://mahasiswa.itenas.ac.id/mahasiswa/AKT298G-matkul-mahasiswa"

# Headers dengan informasi cookie yang diperlukan (isi sesuai kebutuhan)
headers = {
    'Cookie': 'XSRF-TOKEN=eyJpdiI6Iko0ZVNaXC83NzhSanJyZjNyV2xPRWFnPT0iLCJ2YWx1ZSI6IkhCQk1DamRQK3RPRm1UUm5MUWVKRWRJemxzS3RBUGJIR3ltQ082R3FLZUc2OHlXV1UrNElZNEVWdHNPRWhBRm5EclE3aHY1XC9GTnpVWWNmd2JVMHllQT09IiwibWFjIjoiMTRmMDNiOWE4NDEyOTdiMzMyMTczZDVhYmNmOWJiNjYyNzhhNDViZDYxNWM4MmNhYjgzMzQ3OGRkYzI3YmQ3ZCJ9; laravel_session=eyJpdiI6ImFZRU5hRWhTUW15aG1LcFZ4ditWTlE9PSIsInZhbHVlIjoiVnFOZTdwKzlQdkNxOWs4a1owOUsxeWJ6U2FNUEpUWlc2eFdPTGh4YTVDZUVJUEJ5Y3lUOVhaNDU2MnZ6MHNNVEwwUEV5ajd5XC9KOUVkSnB3SE9MMDhBPT0iLCJtYWMiOiJiNzM4NGMwMTgxMmZhMmQyYWM3ZjFhMmFkZmNjNTAxYThlMTI2Zjk3YmE1MWY5ZTJiMTlmYTAzNzQ2NDliZmI2In0%3D'
}

# Nilai kelas dan mata kuliah yang akan diperiksa
target_kelas = "DD"
target_matakuliah = "KEWIRAUSAHAAN"

# Inisialisasi list untuk menyimpan hasil
results = []

# Loop melalui range nimhs dari 152021167 hingga 152021168 (sesuaikan sesuai kebutuhan)
for nimhs in range(152021167, 152021169):
    # Membuat parameter dengan nimhs yang berbeda di setiap iterasi
    params = {
        'sortKode': '',
        'sortHari': '',
        'kode': 'R',
        'thsms': '20241',
        'nimhs': str(nimhs)
    }

    try:
        # Melakukan request GET dengan URL dan parameter
        response = requests.get(base_url, headers=headers, params=params)

        # Memeriksa apakah respons berhasil
        if response.status_code == 200:
            # Mengubah response JSON ke bentuk dictionary
            data = response.json()

            # Variabel untuk melacak status
            matched = False

            # Mengecek setiap entry di dalam data untuk validasi kelas dan mata kuliah
            for item in data:
                matakuliah = item.get('Disp_Matakuliah', '').strip().lower()
                kelas = item.get('Disp_Kelas', '').strip().lower()

                # Debugging: Cetak matakuliah dan kelas yang diproses
                print(f"Processing NIM {nimhs}: Matakuliah='{matakuliah}', Kelas='{kelas}'")

                # Mengecek jika matakuliah dan kelas sesuai dengan target
                if matakuliah == target_matakuliah.lower() and kelas == target_kelas.lower():
                    hari = item.get('Disp_Hari', '').strip()
                    jam = item.get('Disp_Jam', '').strip()

                    # Debugging: Cetak hari dan jam yang ditemukan
                    print(f"  --> Match Found: Hari='{hari}', Jam='{jam}'")

                    results.append({
                        'NIM': nimhs,
                        'Matakuliah': item.get('Disp_Matakuliah', '').strip(),
                        'Kelas': item.get('Disp_Kelas', '').strip(),
                        'Hari': hari,
                        'Jam': jam,
                        'Status': 'Memiliki'
                    })
                    matched = True

            # Jika tidak ada matakuliah yang sesuai, tambahkan status 'Aman'
            if not matched:
                print(f"  --> No matching course found for NIM {nimhs}. Status: Aman")
                results.append({
                    'NIM': nimhs,
                    'Matakuliah': target_matakuliah,
                    'Kelas': target_kelas,
                    'Hari': '',
                    'Jam': '',
                    'Status': 'Aman'
                })

        else:
            print(f"Request failed for NIM: {nimhs} with status code: {response.status_code}")
            results.append({
                'NIM': nimhs,
                'Matakuliah': target_matakuliah,
                'Kelas': target_kelas,
                'Hari': '',
                'Jam': '',
                'Status': f"Request failed with status code {response.status_code}"
            })

    except Exception as e:
        print(f"An error occurred for NIM: {nimhs} - {e}")
        results.append({
            'NIM': nimhs,
            'Matakuliah': target_matakuliah,
            'Kelas': target_kelas,
            'Hari': '',
            'Jam': '',
            'Status': f"Error: {e}"
        })

    # Jeda 1 detik antara request untuk menghindari beban server
    time.sleep(1)

# Membuat DataFrame dari hasil
df = pd.DataFrame(results)

# Menampilkan DataFrame untuk verifikasi sebelum menyimpan
print("\n=== Data yang Diperoleh ===")
print(df)

# Menyimpan DataFrame ke file Excel
excel_file = 'jadwal_mahasiswa.xlsx'
df.to_excel(excel_file, index=False)
print(f"\nData berhasil disimpan ke {excel_file}")
