import pdfplumber
import re
import pandas as pd
import os



# Function to read PDF using pdfplumber and return extracted text
def read_pdf_with_pdfplumber(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        return ''.join(page.extract_text() for page in pdf.pages)


def append_or_write_to_excel(df, filename):
    # Get the directory of the current script
    current_directory = os.path.dirname(os.path.abspath(__file__))
    
    # Combine the directory with the desired filename
    path = os.path.join(current_directory, filename)
    
    if os.path.exists(path):
        # If the file exists, read it
        existing_data = pd.read_excel(path)
        # Append new data
        df = pd.concat([existing_data, df], ignore_index=True)
    
    # Save the updated data or write new file if it doesn't exist
    df.to_excel(path, index=False)

 
# Step 1: Splitting the pdf text into sections
def split_into_sections(text):
    sections = {}
    for i in range(len(headers) - 1):
        start = text.find(headers[i]) + len(headers[i])
        end = text.find(headers[i+1])
        sections[headers[i]] = text[start:end].strip()
    sections[headers[-1]] = text[end + len(headers[-1]):].strip()
    return sections

def get_checked_option(text, stop_pattern=None):
    values = text.split()
    checked_option = None
    for i, val in enumerate(values):
        if val == "X" and i + 1 < len(values):
            checked_option = values[i + 1]
            break
    """ if stop_pattern:
        if re.search(stop_pattern, checked_option):
            checked_option = None """
    return checked_option if checked_option else "SEÇİLMEMİŞ"


def extract_data_from_sections(sections):
    data = {}
    data.update(extract_from_section_1(sections[headers[0]]))
    data.update(extract_from_section_2(sections[headers[1]]))
    data.update(extract_from_section_3(sections[headers[2]]))
    data.update(extract_from_section_4(sections[headers[3]]))
    return data

# Step 2: Extract key-values from sections
def extract_from_section_1(section):
    qr, rg, tar, last, kimlik_no = ("","","","","")
    try:
        # Search for UUID pattern
        qr_search = re.search(r"(\w{8}-\w{4}-\w{4}-\w{4}-\w{12})", section)

        # If the UUID pattern appears before the "RG, TAR, LAST" pattern, use the first format.
        if qr_search and section.index(qr_search.group(1)) < section.index(" - "):
            qr = qr_search.group(1)
            kimlik_no = section.split("\n")[2].strip()
        else:     
            qr = section.split("\n")[2].strip()
            kimlik_no = section.split("\n")[0].strip()
        
        rg_tar_line = section.split("\n")[1]
        rg = rg_tar_line.split("-")[0].strip()
        tar = re.search(r"(\d{2}/\d{2}/\d{4})", rg_tar_line).group(1)
        
        last_l = rg_tar_line.split(" ")
        last = last_l[-2] if last_l[-1] == "-" else last_l[-1]
        
    except Exception as e:
        print('Error :', e )
    
    return {"QR": qr, "RAPOR NO": rg, "MUAYENE KONTROL TARİHİ": tar, "KONTROL TÜRÜ": last, "KİMLİK NO": kimlik_no}

def extract_key_value(section, key, next_key=None):
    
    if key == "ADRES":
        section = collapse_spaces_for_adres(section)
        
    if next_key:
        pattern = re.escape(key) + r"\s*:\s*(.*?)(?=" + re.escape(next_key) + "|\n|$)"
    else:
        pattern = re.escape(key) + r"\s*:\s*(.*)"
    match = re.search(pattern, section)
    if match and match.group(1):
        return match.group(1).strip()
    else:
        return "-"  # Default value when not found


def collapse_spaces_for_adres(text):   
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'A\s+D\s+R\s+E\s+S', 'ADRES', text)
    return text

def extract_from_section_2(section):
    keys = ["ASANSÖR SERİ NO", "MAK. MOTOR SERİ NO",
            "BEYAN YÜKÜ (kg)", "KAT VE DURAK SAYISI",
            "STANDARD/STANDARDLAR",  "ADRES" ] # "BEYAN HIZI (m/sn)",
    for i, key in enumerate(keys):
        next_key =  keys[i+1] if i < len(keys)-1 else None
        val = extract_key_value(section, key, next_key)
        # Ensure to remove spaces between letters at adress value. 
        val = re.sub(r'^(.*?)ADA-PARSEL', lambda m: m.group(1).replace(' ', '') + 'ADA-PARSEL', val) if key == "ADRES" else val
        val = ' / '.join(val.split()) if key == "KAT VE DURAK SAYISI" else val
        val = val.replace('kg', 'kg / ' ) if "BEYAN YÜKÜ" in key else val 
        
        if key == "ADRES":
            data["ASANSÖRÜN ADRESİ"] = val 
        else:
            data[key] = val
            
        if val =="-" and key == "ADRES":
            print(key)
            print(section)
                
            
   # Specific handling for "ASANSÖR SERİ NO"
    asansor_seri_no_pattern = r"ASANSÖR SERİ NO\s*:\s*(.*?)(?=\s*MAK\. MOTOR SERİ NO|$|\n)"
    match = re.search(asansor_seri_no_pattern, section)
    if match and match.group(1):
        data["ASANSÖR SERİ NO"] = match.group(1).strip()
    else:
        data["ASANSÖR SERİ NO"] = "-"  # Default value when not found

    # Specific handling for "MONTAJ YILI"
    montaj_yili_pattern = r"MONTAJ YILI\s*:\s*(\d{4})"
    match = re.search(montaj_yili_pattern, section)
    if match:
        data["MONTAJ YILI"] = match.group(1)
    else:
        data["MONTAJ YILI"] = "Bilgi Yok" 
        print("MONTAJ YILI")
        print(section)
    

    # Specific handling for "SEYİR MESAFESİ (m)"
    seyir_mesafesi_pattern = r"SEYİR MESAFESİ\s*(\(m\))?\s*:\s*(.*?)(\n|$)"
    match = re.search(seyir_mesafesi_pattern, section)
    if match and match.group(2):
        data["SEYİR MESAFESİ (m)"] = match.group(2).strip() + " metre"
    else:
        data["SEYİR MESAFESİ (m)"] = "Bilgi Yok"  # Default value when not found
        print('SEYİR MESAFESİ')
        print(section)

    # Specific handling for "ASANSÖR CİNSİ"
    asansor_cinsi_pattern = r"ASANSÖR CİNSİ\s*:\s*(X\s+)?(İNSAN)?\s*(X\s+)?(YÜK)?\s*(X\s+)?(İNSAN VE YÜK)?"
    match = re.search(asansor_cinsi_pattern, section)
    
    if match:
        if match.group(2) and match.group(1):
            data["ASANSÖR CİNSİ"] = "İNSAN"
        elif match.group(4) and match.group(3):
            data["ASANSÖR CİNSİ"] = "YÜK"
        elif match.group(6) and match.group(5):
            data["ASANSÖR CİNSİ"] = "İNSAN VE YÜK"
        else:
            data["ASANSÖR CİNSİ"] = "SEÇİLMEMİŞ"
            print("ASANSÖR CİNSİ")
            print(section)
            
    # Handling for "ASANSÖR TİPİ"
    asansor_tipi_pattern = r"ASANSÖR TİPİ\s*:\s*(X\s*)?HİDROLİK\s*(X\s*)?ELEKTRİKLİ"
    match = re.search(asansor_tipi_pattern, section)
    if match:
        if match.group(1):  # If the first X is captured
            data["ASANSÖR TİPİ"] = "HİDROLİK"
        elif match.group(2):  # If the second X is captured
            data["ASANSÖR TİPİ"] = "ELEKTRİKLİ"
        else:
            data["ASANSÖR TİPİ"] = "SEÇİLMEMİŞ"
            print("ASANSÖR TİPİ")
            print(section)
               
    # Specific handling for "BEYAN HIZI (m/sn)"
    # beyan_hizi_pattern = r"BEYAN HIZI \(m/sn\)\s*:\s*([\d,]+(\s+[\d,]+)*\s+(X\s+)?[\d,]+(\s+[\d,]+)*\s+DİĞER)"
    beyan_hizi_pattern = r"BEYAN HIZI( \(m/sn\))?\s*:\s*([\s\S]+?(?=\n|$))"

    match = re.search(beyan_hizi_pattern, section)
    if match:
        data["BEYAN HIZI (m/sn)"] = get_checked_option(match.group(2), r"\d+,\d+") + " m/sn"
    else:
        print(section)

    return data

def extract_from_section_3(section):
    keys = ["ADI VE SOYADI", "ADRESİ", "TELEFON NUMARASI", "E-POSTA ADRESİ"]
    data = {}
    for i, key in enumerate(keys):
        if i < len(keys) - 1:
            pattern = re.escape(key) + r"\s*:([\s\S]*?)(?=" + re.escape(keys[i+1]) + ")"
        else:
            pattern = re.escape(key) + r"\s*:(.*)"
        match = re.search(pattern, section)
        if match:
            data['BİNA SORUMLUSU ' + key] = match.group(1).strip()

    permit_search = re.search(r"(PERİYODİK KONTROLE İZİN VERİLDİ\s*:[\s\S]*?)(?=(PERİYODİK KONTROLE İZİN VERİLMEDİ|$))", section)
    if permit_search and "X" in permit_search.group(1):
        data["PERİYODİK KONTROLE İZİN VERİLDİ"] = "EVET"
    else:
        data["PERİYODİK KONTROLE İZİN VERİLDİ"] = "HAYIR"
        
       # Specific handling for "PERİYODİK KONTROLE İZİN VERİLDİ"
    permit_search = re.search(r"PERİYODİK KONTROLE İZİN VERİLDİ\s*:\s*([\s\S]*?)(?=PERİYODİK KONTROLE İZİN VERİLMEDİ|$)", section)
    # Specific handling for "PERİYODİK KONTROLE İZİN VERİLDİ"
    if "X PERİYODİK KONTROLE İZİN VERİLDİ" in section:
        data["PERİYODİK KONTROLE İZİN VERİLDİ"] = "EVET"
    else:
        data["PERİYODİK KONTROLE İZİN VERİLDİ"] = "HAYIR"

    return data



def extract_from_section_4(section):
    pattern = r"ÜNVAN\s*:\s*([\s\S]*?)(?=ADRES\s*:)"
    match = re.search(pattern, section)
    if match:
        unvan = match.group(1).strip()
    else:
        unvan = "Not Found"
    return {"YETKİLİ SERVİS FİRMASI ": unvan}

headers = [
    "ASANSÖR PERİYODİK/TAKİP KONTROL RAPORU",
    "ASANSÖRE İLİŞKİN BİLGİLER",
    "BİNA SORUMLUSUNA İLİŞKİN BİLGİLER",
    "YETKİLİ SERVİSE İLİŞKİN BİLGİ VE BELGELER",
]


if __name__ == "__main__":
    
    # Remove the existing extracted_data.xlsx file if it exists
    """ output_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "extracted_data.xlsx")
    if os.path.exists(output_file_path):
        os.remove(output_file_path) """
    
    # Directory containing all the PDFs
    pdf_directory = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'pdfs')

    # Ensure the directory exists
    if not os.path.exists(pdf_directory):
        print(f"HATA!: Böyle bir klasör yok ! {pdf_directory}")
        exit()
            
    # Loop through each PDF in the directory
    for pdf_file in os.listdir(pdf_directory)[:]:
        if pdf_file.endswith(".pdf"):
            pdf_path = os.path.join(pdf_directory, pdf_file)
            
            pdf_text = read_pdf_with_pdfplumber(pdf_path)
            print(pdf_file, '_'*50)
            sections = split_into_sections(pdf_text)
            data = {}
            
            # Update your data dictionary
            data["TESİS / APARTMAN ADI"] = pdf_file.split('R.')[0]
            
            data.update(extract_from_section_1(sections[headers[0]]))
            data.update(extract_from_section_2(sections[headers[1]]))
            data.update(extract_from_section_3(sections[headers[2]]))
            data.update(extract_from_section_4(sections[headers[3]]))

            # Append to the Excel file (or create it if it doesn't exist)
            df = pd.DataFrame([data])
            append_or_write_to_excel(df, "extracted_data.xlsx")

    print("İşlem Tamamlandı...")