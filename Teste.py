import cv2
import easyocr
import pandas as pd
import re
import time
import os

reader = easyocr.Reader(['pt'])

# -------- LIMPEZA --------

def limpar_telefone(numero):
    numero = re.sub(r'[^\d+]', '', numero)
    return numero

def limpar_email(email):
    email = email.replace(" ", "")
    email = email.strip().lower()
    return email


def extrair_dados_pt(texto_lista, nome_empresa, morada_manual, servico):

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

    termos_morada = ['rua', 'av', 'avenida', 'travessa', 'zona', 'industrial', 'lote', 'edif', 'praça', 'apartado', 'urbanização']
    linhas_morada_detectadas = []

    # REGEX MELHORADAS
    re_fixo = r'(\+351\s?)?2\d{2}[\s.-]?\d{3}[\s.-]?\d{3}'
    re_movel = r'(\+351\s?)?9\d{2}[\s.-]?\d{3}[\s.-]?\d{3}'
    re_cp = r'(\d{4}-\d{3})'
    re_email = r'\b[A-Za-z0-9._%+-]+\s*@\s*[A-Za-z0-9.-]+\s*\.\s*[A-Za-z]{2,}\b'
    re_nif = r'\b([1256789]\d{8})\b'
    re_site = r'\b(www\.[\w\.-]+|https?://[\w\.-]+|[\w\.-]+\.(pt|com|net|org))\b'

    for item in texto_lista:

        item_clean = item.strip()
        item_lower = item_clean.lower()

        # -------- EMAIL --------
        email_match = re.search(re_email, item_clean)

        if email_match:
            dados["Email"] = limpar_email(email_match.group(0))
            continue


        # -------- TELEFONES --------

        fixo_match = re.search(re_fixo, item_clean)
        if fixo_match:
            dados["telefone_fixo"] = limpar_telefone(fixo_match.group(0))

        movel_match = re.search(re_movel, item_clean)
        if movel_match:
            dados["Telemóvel/Telefone"] = limpar_telefone(movel_match.group(0))


        # -------- CODIGO POSTAL --------

        cp_match = re.search(re_cp, item_clean)

        if cp_match:
            dados["Código-Postal"] = cp_match.group(1)

            if item_clean not in linhas_morada_detectadas:
                linhas_morada_detectadas.append(item_clean)


        # -------- NIF --------

        nif_match = re.search(re_nif, item_clean)

        if nif_match:
            nif_val = nif_match.group(1)

            tel_clean = re.sub(r'\D', '', dados["Telemóvel/Telefone"])

            if nif_val not in tel_clean:
                dados["NIF/NIPC"] = nif_val


        # -------- SITE --------

        site_match = re.search(re_site, item_clean, re.IGNORECASE)

        if site_match:

            s = site_match.group(0).lower()

            if not re.match(r'https?://', s):
                s = 'http://' + s

            dados["Site"] = s


        # -------- MORADA --------

        if any(termo in item_lower for termo in termos_morada):

            if item_clean not in linhas_morada_detectadas:
                linhas_morada_detectadas.append(item_clean)


    if not dados["morada"] and linhas_morada_detectadas:

        dados["morada"] = " - ".join(linhas_morada_detectadas)

    return dados



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


        gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)

        img_pre = cv2.resize(gray, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)


        result = reader.readtext(img_pre, detail=0)


        if not result:
            print("Nenhum texto detectado. Tente aproximar mais o cartão.")
            continue


        # JUNTA TEXTO OCR (AJUDA EMAILS PARTIDOS)
        texto_completo = " ".join(result)
        result.append(texto_completo)


        info = extrair_dados_pt(result, nome_empresa, morada_manual, servico)


        print(f"Dados Capturados: {info}")


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