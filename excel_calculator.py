import pandas as pd
import json

# Fungsi untuk membaca data dari file Excel
def read_excel(file_path, sheet_name=0, keyword=None):
    try:
        # Membaca file Excel
        data = pd.read_excel(file_path, sheet_name=sheet_name)

        # Pastikan kolom Timestamp diubah ke datetime
        if "Timestamp" in data.columns:
            data["Timestamp"] = pd.to_datetime(data["Timestamp"], errors='coerce')

        # Filter berdasarkan kata kunci
        if keyword and "Nama Tim" in data.columns:
            data = data[data["Nama Tim"] == keyword]

        return data
    except Exception as e:
        print(f"Terjadi kesalahan saat membaca file: {e}")
        return None

# Fungsi untuk menghitung data dan mengelompokkan berdasarkan tanggal
def calculate_keyword_data(data):
    try:
        # Pastikan kolom Timestamp tersedia
        if "Timestamp" in data.columns:
            # Mengelompokkan data berdasarkan tanggal
            data['Date'] = data['Timestamp'].dt.date  # Extract hanya tanggalnya

            # Kelompokkan data berdasarkan tanggal
            grouped_data = data.groupby('Date')

            # Menyusun hasil untuk setiap grup
            result = {}
            for date, group in grouped_data:
                result[str(date)] = {
                    "Total Occurrences": len(group),
                    "Entry Times": group['Timestamp'].dt.strftime('%Y-%m-%d %H:%M:%S').tolist(),
                }
            return result
        else:
            print("Kolom 'Timestamp' tidak ditemukan.")
            return None
    except Exception as e:
        print(f"Terjadi kesalahan saat menghitung data: {e}")
        return None

# Fungsi untuk menyimpan hasil ke file JSON
def save_to_json(data, output_file):
    try:
        with open(output_file, "w", encoding="utf-8") as file:
            json.dump(data, file, indent=4, ensure_ascii=False)
        print(f"Hasil telah disimpan di file: {output_file}")
    except Exception as e:
        print(f"Terjadi kesalahan saat menyimpan file: {e}")

# Main program
if __name__ == "__main__":
    # Path ke file Excel (ganti dengan path file Anda)
    file_path = "KUESIONER VALIDASI MARKET WMK 2024 (Responses).xlsx"
    output_file = "output_filtered_by_date.json"

    # Kata kunci untuk filter
    keyword = "21A ProtectSphere: Red Force"

    # Membaca data
    print("Membaca data...")
    data = read_excel(file_path, keyword=keyword)

    if data is not None:
        print("Data berhasil dibaca dan difilter!")
        print("Data Preview:")
        print(data.head())  # Menampilkan 5 baris pertama

        # Menghitung statistik dan mengelompokkan berdasarkan tanggal
        print("\nMenghitung statistik data dan mengelompokkan berdasarkan tanggal...")
        stats = calculate_keyword_data(data)

        if stats:
            print("\nHasil Perhitungan:")
            print(json.dumps(stats, indent=4, ensure_ascii=False))  # Menampilkan hasil di terminal

            # Menyimpan ke file JSON
            save_to_json(stats, output_file)
