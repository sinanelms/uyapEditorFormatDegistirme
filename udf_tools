import os
import zipfile
from docx import Document
import xml.etree.ElementTree as ET

def udf_to_docx(udf_file_path):
    """
    UYAP UDF dosyasını (ZIP tabanlı) DOCX'e dönüştürür
    """
    if not os.path.exists(udf_file_path):
        print(f"Hata: UDF dosyası bulunamadı: {udf_file_path}")
        return
    
    file_name, _ = os.path.splitext(os.path.basename(udf_file_path))
    docx_file_path = os.path.join(os.path.dirname(udf_file_path), f"{file_name}.docx")
    extract_dir = os.path.join(os.path.dirname(udf_file_path), f"{file_name}_extracted")
    
    try:
        # UDF dosyasını ZIP olarak aç
        print("UDF dosyası ZIP arşivi olarak açılıyor...")
        with zipfile.ZipFile(udf_file_path, 'r') as zip_ref:
            # İçeriği listele
            print("\nArşiv içeriği:")
            for file_info in zip_ref.filelist:
                print(f"  - {file_info.filename} ({file_info.file_size} bytes)")
            
            # Tüm içeriği çıkart
            zip_ref.extractall(extract_dir)
            print(f"\nDosyalar çıkartıldı: {extract_dir}")
        
        # Çıkartılan dosyaları kontrol et
        print("\n--- Çıkartılan dosyalar analiz ediliyor ---")
        
        # Word belge içeriğini bul (genellikle document.xml veya content.xml)
        content_files = []
        for root, dirs, files in os.walk(extract_dir):
            for file in files:
                if file.endswith('.xml') or file.endswith('.txt'):
                    full_path = os.path.join(root, file)
                    content_files.append(full_path)
                    print(f"\nBulunan dosya: {file}")
                    
                    # İlk 500 karakteri göster
                    try:
                        with open(full_path, 'r', encoding='utf-8', errors='ignore') as f:
                            content_preview = f.read(500)
                            print(f"İçerik önizleme: {content_preview[:200]}...")
                    except:
                        pass
        
        # DOCX oluştur
        doc = Document()
        doc.add_heading('UDF Dosyası İçeriği', 0)
        
        # Metin içeriğini topla
        for content_file in content_files:
            try:
                with open(content_file, 'r', encoding='utf-8', errors='ignore') as f:
                    content = f.read()
                    
                    # XML ise parse et
                    if content_file.endswith('.xml'):
                        try:
                            # Basit metin çıkar
                            root = ET.fromstring(content)
                            text = ' '.join(root.itertext())
                            if text.strip():
                                doc.add_paragraph(text.strip())
                        except:
                            # XML parse edilemezse düz metin olarak ekle
                            doc.add_paragraph(content)
                    else:
                        doc.add_paragraph(content)
            except Exception as e:
                print(f"Hata ({content_file}): {e}")
        
        # DOCX kaydet
        doc.save(docx_file_path)
        print(f"\n✓ Başarıyla dönüştürüldü: {docx_file_path}")
        
        # Temizlik (isteğe bağlı)
        # import shutil
        # shutil.rmtree(extract_dir)
        
    except zipfile.BadZipFile:
        print("Hata: Dosya geçerli bir ZIP arşivi değil")
    except Exception as e:
        print(f"Hata: {e}")
        import traceback
        traceback.print_exc()

def udf_to_txt(udf_dosya_yolu):
    """
    UDF dosyasını TXT'ye dönüştürür
    """
    if not os.path.exists(udf_dosya_yolu):
        print(f"Hata: UDF dosyası bulunamadı: {udf_dosya_yolu}")
        return
    
    file_name, _ = os.path.splitext(os.path.basename(udf_dosya_yolu))
    txt_dosya_yolu = os.path.join(os.path.dirname(udf_dosya_yolu), f"{file_name}.txt")
    extract_dir = os.path.join(os.path.dirname(udf_dosya_yolu), f"{file_name}_extracted")
    
    try:
        # ZIP olarak aç ve içeriği çıkart
        with zipfile.ZipFile(udf_dosya_yolu, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        
        # Tüm metin içeriğini topla
        all_text = []
        for root, dirs, files in os.walk(extract_dir):
            for file in files:
                if file.endswith(('.xml', '.txt')):
                    full_path = os.path.join(root, file)
                    try:
                        with open(full_path, 'r', encoding='utf-8', errors='ignore') as f:
                            content = f.read()
                            
                            if file.endswith('.xml'):
                                try:
                                    root_elem = ET.fromstring(content)
                                    text = ' '.join(root_elem.itertext())
                                    all_text.append(text.strip())
                                except:
                                    all_text.append(content)
                            else:
                                all_text.append(content)
                    except:
                        pass
        
        # TXT dosyasına yaz
        with open(txt_dosya_yolu, 'w', encoding='utf-8') as txt_file:
            txt_file.write('\n\n'.join(all_text))
        
        print(f"✓ TXT dosyası oluşturuldu: {txt_dosya_yolu}")
        
    except Exception as e:
        print(f"Hata: {e}")

# --- KULLANIM ---
script_dir = os.path.dirname(os.path.abspath(__file__))
udf_dosyasi = os.path.join(script_dir, "a.udf")

print("=== UDF'den DOCX'e Dönüştürme ===")
udf_to_docx(udf_dosyasi)

print("\n=== UDF'den TXT'ye Dönüştürme ===")
udf_to_txt(udf_dosyasi)
