import cv2
import easyocr
import pandas as pd
import re
import time
import os

# 1. Configura EasyOCR para Português
# O modelo será carregado na primeira execução
reader = easyocr.Reader(['pt'])

def extrair_dados_pt(texto_lista, nome_empresa, morada_manual,servico):
    """Filtra dados seguindo padrões de Portugal (Fixo, Móvel e Morada)"""
    dados = {
        "Empresa/Nome": nome_empresa,
        "morada": morada_manual,
        "Código-Postal": "",
        "NIF/NIPC": "",
        "servico a oferecer": servico,
        "telefone_fixo": "",                  
        "Telemóvel/Telefone": "",
        "Email": "",
        "Site": ""

    }

    # Palavras-chave para identificar linhas de morada
    termos_morada = ['rua', 'av', 'avenida', 'travessa', 'zona', 'industrial', 'lote', 'edif', 'praça', 'apartado', 'urbanização']
    linhas_morada_detectadas = []

    # Expressões Regulares (RegEx)
    # Suporta: 2XXXXXXXX, +351 2XXXXXXXX, 9XXXXXXXX, +351 9XXXXXXXX
    re_fixo = r'(\+351[\s.-]?)?(2\d{8}|2\d{1}[\s.-]?\d{3,4}[\s.-]?\d{3,4})'
    re_movel = r'(\+351[\s.-]?)?([39]\d{8}|[39]\d{2}[\s.-]?\d{3}[\s.-]?\d{3})'
    re_cp = r'(\d{4}-\d{3})'
    re_email = r'[\w\.-]+@[\w\.-]+'
    re_nif = r'\b([1256789]\d{8})\b'
    re_site = r'\b(www\.[\w\.-]+|https?://[\w\.-]+|[\w\.-]+\.(pt|com|net|org))\b'

    for item in texto_lista:
        item_clean = item.strip()
        item_lower = item_clean.lower()

        # --- 1. Email ---
        if re.search(re_email, item_clean):
            dados["Email"] = item_clean.lower()
            continue # Se é email, ignoramos para morada

        # --- 2. Telefones (Fixo e Móvel) ---
        fixo_match = re.search(re_fixo, item_clean)
        if fixo_match:
            dados["telefone_fixo"] = fixo_match.group(0)
        
        movel_match = re.search(re_movel, item_clean)
        if movel_match:
            dados["Telemóvel/Telefone"] = movel_match.group(0)

        # --- 3. Código Postal e Morada na mesma linha ---
        cp_match = re.search(re_cp, item_clean)
        if cp_match:
            dados["Código-Postal"] = cp_match.group(1)
            # Se tem CP, a linha quase de certeza é a morada
            if item_clean not in linhas_morada_detectadas:
                linhas_morada_detectadas.append(item_clean)

        # --- 4. NIF ---
        nif_match = re.search(re_nif, item_clean)
        if nif_match:
            nif_val = nif_match.group(1)
            # Evita confundir NIF com partes do telefone
            tel_clean = re.sub(r'\D', '', dados["Telemóvel/Telefone"])
            if nif_val not in tel_clean:
                dados["NIF/NIPC"] = nif_val
        
        # --- 5. Site ---
        site_match = re.search(re_site, item_clean, re.IGNORECASE)
        if site_match:  
            s = site_match.group(0).lower()
            if not re.match(r'https?://', s):
                s = 'http://' + s
            dados["Site"] = s

       

        if any(termo in item_lower for termo in termos_morada):
            if item_clean not in linhas_morada_detectadas:
                linhas_morada_detectadas.append(item_clean)

    # Se a morada manual estiver vazia, usamos o que o OCR encontrou
    if not dados["morada"] and linhas_morada_detectadas:
        # Junta as linhas (Rua + CP/Localidade)
        dados["morada"] = " - ".join(linhas_morada_detectadas)

    return dados


# 2. Conecta à câmera (DroidCam ou Webcam)
cap = cv2.VideoCapture(0) 
if not cap.isOpened():
    print("Erro: não conseguiu acessar a câmera.")
    exit()

print("Câmara aberta. Pressione 's' para digitalizar | 'q' para sair")
arquivo_excel = 'contactos_empresas_pt.xlsx'

while True:
    ret, frame = cap.read()
    if not ret:
        break

    cv2.imshow('Leitor de Cartoes PT', frame)
    key = cv2.waitKey(1) & 0xFF

    if key == ord('s'):
        print("\n--- Processando Cartão ---")
        nome_empresa = input("Digite o nome da empresa (ou deixe vazio): ")
        servico = input("Digite o serviço a oferecer: ")
        morada_manual = input("Digite a morada (ou deixe vazio para detecção automática): ")

        # PRÉ-PROCESSAMENTO PARA LETRAS PEQUENAS 
        # Converte para cinza para reduzir ruído
        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
        # Amplia a imagem em 2x para ajudar o OCR a ver letras pequenas
        img_pre = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
        
        # OCR na imagem melhorada
        result = reader.readtext(img_pre, detail=0)
        
        if not result:
            print("Nenhum texto detectado. Tente aproximar mais o cartão.")
            continue

        # Extrai os dados
        info = extrair_dados_pt(result, nome_empresa, morada_manual,servico)  # Passa o serviço a oferecer para o dicionário
        
        # Mostra o que foi capturado no terminal para conferir
        print(f"Dados Capturados: {info}")

        # Salva no Excel
        df = pd.DataFrame([info])
        try:
            df_old = pd.read_excel(arquivo_excel)
            df_final = pd.concat([df_old, df], ignore_index=True)
        except FileNotFoundError:
            df_final = df

        df_final.to_excel(arquivo_excel, index=False)
        print(f"Sucesso! Registado em {arquivo_excel}")

    elif key == ord('q'):
        print("Saindo...")
        break

cap.release()
cv2.destroyAllWindows()
if os.path.exists(arquivo_excel):
    print(f"A abrir {arquivo_excel}...")
    os.system(f'start {arquivo_excel}')

