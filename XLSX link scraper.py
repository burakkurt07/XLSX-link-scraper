import openpyxl

# Kullanıcıdan dosya yolunu iste
file_path = input("Lütfen Excel dosyasının tam yolunu girin (örn: C:\\Users\\yourusername\Desktop\file.xlsx): ")

try:
    # Excel dosyasını yükle
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active  # İlk sayfa

    # Linkleri toplamak için bir liste
    links = []

    # B sütunundaki linkleri al
    for row in sheet.iter_rows(min_row=1, max_col=2, values_only=False):
        cell = row[1]  # B sütunundaki hücre
        if cell.hyperlink:  # Eğer hücrede link varsa
            links.append(cell.hyperlink.target)  # Link adresini al

    # TXT dosyasına kaydet
    with open("links.txt", "w", encoding="utf-8") as f:
        for link in links:
            f.write(link + "\n")

    print("Linkler 'links.txt' dosyasına başarıyla kaydedildi.")

except FileNotFoundError:
    print(f"Hata: '{file_path}' dosyası bulunamadı. Lütfen dosya yolunu kontrol edin.")
except Exception as e:
    print(f"Beklenmeyen bir hata oluştu: {e}")
