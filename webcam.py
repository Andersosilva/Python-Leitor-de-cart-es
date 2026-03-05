import cv2
import easyocr
import pandas as pd
import re
import time


# 1. Configura EasyOCR para Português

reader = easyocr.Reader(['pt'])

def extrair_dados_pt(texto_lista, nome_empresa, morada):
    """Filtra dados seguindo padrões de Portugal (Fixo e Móvel)"""
    dados = {
        "Empresa/Nome": nome_empresa,
        "morada": morada,
        "telefone_fixo": "",                  
        "Telemóvel/Telefone": "",
        "Email": "",
        "Código-Postal": "",
        "NIF/NIPC": "",
        "Site": ""
    }

    # Expressões Regulares Melhoradas
    # Suporta: 2XXXXXXXX, +351 2XXXXXXXX, 9XXXXXXXX, +351 9XXXXXXXX, etc.
    padrao_fixo = r'(\+351[\s.-]?)?(2\d{8}|2\d{1}[\s.-]?\d{3,4}[\s.-]?\d{3,4})'
    padrao_movel = r'(\+351[\s.-]?)?([39]\d{8}|[39]\d{2}[\s.-]?\d{3}[\s.-]?\d{3})'

    for item in texto_lista:
        item = item.strip()

        # 1. Email
        if re.search(r'[\w\.-]+@[\w\.-]+', item):
            dados["Email"] = item.lower()
            print(f"Email detectado: {dados['Email']}")

        # 2. Telefone Fixo (Prioridade para números que começam com 2)
        fixo_match = re.search(padrao_fixo, item)
        if fixo_match:
            # .group(0) pega a correspondência inteira (incluindo o +351 se existir)
            dados["telefone_fixo"] = fixo_match.group(0)
            print(f"Fixo detectado: {dados['telefone_fixo']}")

        # 3. Telemóvel ou Outros (Começam com 9 ou 3)
        movel_match = re.search(padrao_movel, item)
        if movel_match:
            dados["Telemóvel/Telefone"] = movel_match.group(0)
            print(f"Telemóvel detectado: {dados['Telemóvel/Telefone']}")
        
        # 4. Código Postal
        cp = re.search(r'(\d{4}-\d{3})', item)
        if cp:
            dados["Código-Postal"] = cp.group(1)
            print(f"Código Postal detectado: {dados['Código-Postal']}")
        
        # 5. NIF 
        nif = re.search(r'\b([1256789]\d{8})\b', item)
        if nif:
            nif_val = nif.group(1)
            # Limpeza simples para comparação
            tel_clean = re.sub(r'\D', '', dados["Telemóvel/Telefone"])
            fixo_clean = re.sub(r'\D', '', dados["telefone_fixo"])
            
            if nif_val not in tel_clean and nif_val not in fixo_clean:
                dados["NIF/NIPC"] = nif_val
                print(f"NIF detectado: {nif_val}")
        
        site_match = re.search(r'\b(www\.[\w\.-]+|https?://[\w\.-]+|[\w\.-]+\.(pt|com|net|org))\b', item, re.IGNORECASE)
        if site_match:
            s = site_match.group(0).lower()
            if not re.match(r'https?://', s):
                s = 'http://' + s
            dados["site"] = s
            print(f"Site Reconhecido: {dados['site']}")

    return dados


# 2. Conecta à câmera DroidCam via USB

cap = cv2.VideoCapture(0)  #  0, 1 ou 2 
if not cap.isOpened():
    print("Erro: não conseguiu acessar a câmera ")
    exit()

print("Câmara aberta. Pressione 's' para salvar cartão | 'q' para sair")
arquivo_excel = 'contactos_empresas_pt.xlsx'

#  Loop principal para múltiplos cartões

while True:
    ret, frame = cap.read()
    if not ret:
        print("Falha ao capturar frame. Tentando novamente...")
        continue

    cv2.imshow('Leitor de Cartoes PT', frame)
    key = cv2.waitKey(1) & 0xFF

    if key == ord('s'):  # Salvar cartão
        timestamp = int(time.time())
        filename = f'cartao_{timestamp}.jpg'
        cv2.imwrite(filename, frame)
        print(f"Imagem capturada: {filename}")
        nome_empresa = input("Digite o nome da empresa: ")  # Solicita o nome da empresa para associar ao cartão
        morada = input("Digite a morada da empresa: ")  # Solicita a morada da empresa para associar ao cartão

        # OCR diretamente do frame (não precisa salvar arquivo)
        result = reader.readtext(frame, detail=0)
        if not result:
            print("Nenhum texto detectado. Tente novamente.")
            continue

        # Extrai dados e salva no Excel
        info = extrair_dados_pt(result, nome_empresa,morada)
        df = pd.DataFrame([info])

        try:
            df_old = pd.read_excel(arquivo_excel)
            df_final = pd.concat([df_old, df], ignore_index=True)
        except FileNotFoundError:
            df_final = df

        df_final.to_excel(arquivo_excel, index=False)
        print(f"Cartão registrado com sucesso no arquivo {arquivo_excel}!")

    elif key == ord('q'):  # Sair
        print("Saindo...")
        break

cap.release()
cv2.destroyAllWindows()