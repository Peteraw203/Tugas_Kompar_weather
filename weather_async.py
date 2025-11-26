import asyncio 
import aiohttp 
import pandas as pd 

async def ambil_data(session, lokasi): # Membuat fungsi async yang menerima sesi koneksi dan nama lokasi
    url = f"http://api.weatherapi.com/v1/current.json?key={"242819b7dfe84b4488a34756251809"}&q={lokasi}&aqi=no" # Menyusun URL API berdasarkan lokasi
    async with session.get(url) as response: # Melakukan request GET ke URL API
        data = await response.json() # Menunggu balasan server dan mengubahnya jadi data JSON (Dictionary)
        
        data_json = {
            'Last Update': data['location']['localtime'], # Mengambil waktu update lokal
            'Suhu (°C)': data['current']['temp_c'], # Mengambil suhu celcius
            'Kelembapan (%)': data['current']['humidity'], # Mengambil kelembapan
            'Kondisi Cuaca': data['current']['condition']['text'], # Mengambil teks kondisi cuaca
            'Kecepatan Angin (km/h)': data['current']['wind_kph'], # Mengambil kecepatan angin
            'Arah Angin': data['current']['wind_dir'], # Mengambil arah angin
            'Sinar UV': data['current']['uv'] # Mengambil indeks UV
        }
        # print(f"data {lokasi}: {data_json}")
        # Mengembalikan data yang dipilih dalam bentuk Dictionary 
        return data_json
         
async def main(): # Fungsi utama program
    df = pd.read_excel("excel_weather_async.xlsx") # Membaca file Excel dan menyimpannya ke variabel df
    
    async with aiohttp.ClientSession() as session: # Membuka koneksi internet (session)
        # Membuat daftar tugas (tasks) untuk mengambil data semua kecamatan sekaligus (List Comprehension)
        tasks = [ambil_data(session, kota) for kota in df['Kecamatan']] 
        
        # Menjalankan semua tugas secara bersamaan dan menunggu hasilnya terkumpul di variabel 'hasil'
        hasil = await asyncio.gather(*tasks) 

    # Mengisi setiap kolom Excel dengan mengambil data dari list 'hasil'
    df['Last Update'] = [x['Last Update'] for x in hasil] # Mengisi kolom Last Update
    df['Suhu (°C)'] = [x['Suhu (°C)'] for x in hasil] # Mengisi kolom Suhu
    df['Kelembapan (%)'] = [x['Kelembapan (%)'] for x in hasil] # Mengisi kolom Kelembapan
    df['Kondisi Cuaca'] = [x['Kondisi Cuaca'] for x in hasil] # Mengisi kolom Kondisi Cuaca
    df['Kecepatan Angin (km/h)'] = [x['Kecepatan Angin (km/h)'] for x in hasil] # Mengisi kolom Kecepatan Angin
    df['Arah Angin'] = [x['Arah Angin'] for x in hasil] # Mengisi kolom Arah Angin
    df['Sinar UV'] = [x['Sinar UV'] for x in hasil] # Mengisi kolom Sinar UV

    df.to_excel("excel_weather_async.xlsx", index=False) # Menyimpan kembali perubahan ke file Excel
    print("Selesai! Data berhasil diambil dan disimpan.") # Memberi tahu user program selesai

if __name__ == "__main__": # Mengecek apakah script dijalankan langsung
    asyncio.run(main()) # Menjalankan fungsi main dengan asyncio