# Confer-ncia-NF-S-PDF
Conferência gerada para analisar as notas fiscais encaminhadas pelos colaboradores, no fim trazendo um resultado em Excel da sucessão.
import os
import re
from pathlib import Path

import pandas as pd
import pdfplumber
import pytesseract
from pdf2image import convert_from_path


# =========================
# CONFIGURAÇÕES WINDOWS
# =========================

# Pasta onde estão os PDFs
FOLDER_PATH = Path(r"C:\Users\elias.santos\Desktop\Elias Henrique\TESTE NOTAS FISCAIS 2\RA-01")

# Arquivo de saída Excel
OUTPUT_FILE = FOLDER_PATH / "conferencia_notas_v5.xlsx"

# Caminho do Tesseract no Windows
# Ajuste se necessário
TESSERACT_CMD = r"C:\Users\elias.santos\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"

# Caminho do Poppler no Windows
# Ajuste se necessário
# Exemplo: C:\poppler\Library\bin
POPPLER_PATH = r"C:\poppler\Library\bin"

# Idioma OCR
OCR_LANG = "por"


# Configura o caminho do executável do Tesseract, se existir
if os.path.exists(TESSERACT_CMD):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_CMD


def extract_text_from_pdf(pdf_path: Path) -> str:
    text = ""

    try:
        # 1) Tenta extrair texto nativo do PDF
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"

        # 2) Se o texto veio ruim, tenta OCR
        if not text.strip() or len(text) < 200 or "R$" not in text:
            convert_kwargs = {}
            if os.path.exists(POPPLER_PATH):
                convert_kwargs["poppler_path"] = POPPLER_PATH

            images = convert_from_path(str(pdf_path), **convert_kwargs)

            ocr_text = ""
            for img in images:
                ocr_text += pytesseract.image_to_string(img, lang=OCR_LANG) + "\n"

            # Se OCR trouxe conteúdo, soma ao texto
            if ocr_text.strip():
                text += "\n" + ocr_text

    except Exception as e:
        print(f"[ERRO] Falha ao processar '{pdf_path.name}': {e}")

    return text


def clean_value(val_str: str) -> str:
    if not val_str:
        return "0,00"

    cleaned = re.sub(r"[^0-9,.]", "", val_str)
    if not cleaned:
        return "0,00"

    cleaned = cleaned.strip(",.")
    return cleaned if cleaned else "0,00"


def parse_nf_data(text: str, filename: str) -> dict:
    data = {
        "Arquivo": filename,
        "Número da Nota": "Não encontrado",
        "Data de Emissão": "Não encontrado",
        "Prestador Nome": "Não encontrado",
        "Prestador CNPJ": "Não encontrado",
        "Tomador Nome": "Não encontrado",
        "Tomador CNPJ": "Não encontrado",
        "Item LC / CNAE": "Não encontrado",
        "Valor Bruto": "0,00",
        "Valor Líquido": "0,00",
        "Impostos": "Nenhum destacado",
        "Optante Simples": "Não",
    }

    text_upper = text.upper()

    # Número da Nota
    num_match = re.search(
        r"(?:Número da Nota Fiscal|Número da NFS-e|Número da Nota|Nº da Nota|Nº da NFS-e)\s*[:\-]?\s*(\d+)",
        text,
        re.I,
    )
    if num_match:
        data["Número da Nota"] = num_match.group(1)
    else:
        lines = text.splitlines()
        for line in lines[:15]:
            clean_line = line.strip()
            if re.fullmatch(r"\d{1,8}", clean_line):
                data["Número da Nota"] = clean_line
                break

    # Data de Emissão
    date_match = (
        re.search(
            r"(?:Data de Emissão|Data de Geração|Data de Competência|Emissão em)\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})",
            text,
            re.I,
        )
        or re.search(r"(\d{2}/\d{2}/\d{4})", text)
    )
    if date_match:
        data["Data de Emissão"] = date_match.group(1)

    # CNPJs gerais
    cnpjs = re.findall(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", text)

    # Prestador
    if "PRESTADOR" in text_upper:
        parts = text_upper.split("PRESTADOR")
        if len(parts) > 1:
            section_p = parts[1].split("TOMADOR")[0]

            p_cnpj = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", section_p)
            if p_cnpj:
                data["Prestador CNPJ"] = p_cnpj.group(1)

            p_nome = re.search(
                r"(?:NOME/RAZÃO SOCIAL|RAZÃO SOCIAL|NOME)\s*[:\-]?\s*([^\n]+)",
                section_p,
                re.I,
            )
            if p_nome:
                data["Prestador Nome"] = p_nome.group(1).strip()

    # Tomador
    if "TOMADOR" in text_upper:
        parts = text_upper.split("TOMADOR")
        if len(parts) > 1:
            section_t = (
                parts[1]
                .split("INTERMEDIÁRIO")[0]
                .split("DESCRIÇÃO")[0]
                .split("DISCRIMINAÇÃO")[0]
            )

            t_cnpj = re.search(r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})", section_t)
            if t_cnpj:
                data["Tomador CNPJ"] = t_cnpj.group(1)

            t_nome = re.search(
                r"(?:NOME/RAZÃO SOCIAL|RAZÃO SOCIAL|NOME)\s*[:\-]?\s*([^\n]+)",
                section_t,
                re.I,
            )
            if t_nome:
                data["Tomador Nome"] = t_nome.group(1).strip()

    # Fallback CNPJ
    if data["Prestador CNPJ"] == "Não encontrado" and len(cnpjs) > 0:
        data["Prestador CNPJ"] = cnpjs[0]

    if data["Tomador CNPJ"] == "Não encontrado" and len(cnpjs) > 1:
        data["Tomador CNPJ"] = cnpjs[1]

    # CNAE / Item LC
    cnae_match = re.search(
        r"(?:Cód\.?\s*CNAE|Item da LC116/2003|Código do Serviço|CNAE|Item da LC)\s*[:\-]?\s*([\d.]+)",
        text,
        re.I,
    )
    if cnae_match:
        data["Item LC / CNAE"] = cnae_match.group(1)

    # Valor Bruto
    bruto_patterns = [
        r"VALOR TOTAL DO SERVIÇO\s*=?\s*R\$\s*([\d.,]+)",
        r"VL\.?\s*TOTAL DOS SERVIÇOS\s*R\$\s*([\d.,]+)",
        r"VALOR TOTAL DA NFS-E\s*R\$\s*([\d.,]+)",
        r"VALOR DO SERVIÇO\s*R\$\s*([\d.,]+)",
        r"TOTAL\s*[:\-]?\s*R\$\s*([\d.,]+)",
        r"VALOR TOTAL\s*R\$\s*([\d.,]+)",
        r"VALOR BRUTO\s*R\$\s*([\d.,]+)",
    ]

    for pattern in bruto_patterns:
        m = re.search(pattern, text, re.I)
        if m:
            data["Valor Bruto"] = clean_value(m.group(1))
            break

    # Valor Líquido
    liq_patterns = [
        r"VL\.?\s*LÍQUIDO DA NOTA FISCAL\s*R\$\s*([\d.,]+)",
        r"VALOR LÍQUIDO DA NFS-E\s*R\$\s*([\d.,]+)",
        r"VALOR LÍQUIDO\s*R\$\s*([\d.,]+)",
        r"LÍQUIDO\s*R\$\s*([\d.,]+)",
        r"VALOR LÍQUIDO DA NOTA\s*R\$\s*([\d.,]+)",
    ]

    for pattern in liq_patterns:
        m = re.search(pattern, text, re.I)
        if m:
            data["Valor Líquido"] = clean_value(m.group(1))
            break

    # Fallback para bruto
    if data["Valor Bruto"] == "0,00":
        all_values = re.findall(r"R\$\s*([1-9][\d.,]*)", text)
        if all_values:
            data["Valor Bruto"] = clean_value(all_values[0])

    # Fallback líquido = bruto
    if data["Valor Líquido"] == "0,00" and data["Valor Bruto"] != "0,00":
        data["Valor Líquido"] = data["Valor Bruto"]

    # Impostos
    impostos = []
    for imp in ["PIS", "COFINS", "IRRF", "CSLL", "ISS", "ISSQN"]:
        imp_match = re.search(rf"{imp}\s*[:\-]?\s*R\$\s*([1-9][\d.,]*)", text, re.I)
        if imp_match:
            impostos.append(f"{imp}: {clean_value(imp_match.group(1))}")

    if impostos:
        data["Impostos"] = ", ".join(impostos)

    # Simples Nacional
    if re.search(
        r"OPTANTE PELO SIMPLES NACIONAL|OPTANTE\s*-\s*MICROEMPRESA|SIMPLES NACIONAL|OPTANTE PELO SIMPLES",
        text,
        re.I,
    ):
        data["Optante Simples"] = "Sim"

    return data


def check_dependencies():
    problems = []

    if not FOLDER_PATH.exists():
        problems.append(f"Pasta não encontrada: {FOLDER_PATH}")

    if not os.path.exists(TESSERACT_CMD):
        problems.append(
            "Tesseract não encontrado no caminho configurado. "
            f"Verifique: {TESSERACT_CMD}"
        )

    if not os.path.exists(POPPLER_PATH):
        problems.append(
            "Poppler não encontrado no caminho configurado. "
            f"Verifique: {POPPLER_PATH}"
        )

    return problems


def main():
    print("Iniciando conferência de notas fiscais...")

    dependency_problems = check_dependencies()
    if dependency_problems:
        print("\n[ATENÇÃO] Foram encontrados problemas de configuração:")
        for p in dependency_problems:
            print(f" - {p}")
        print("\nO script tentará continuar, mas OCR/PDF imagem pode falhar.\n")

    files = [f for f in FOLDER_PATH.iterdir() if f.is_file() and f.suffix.lower() == ".pdf"]

    if not files:
        print(f"Nenhum PDF encontrado em: {FOLDER_PATH}")
        return

    results = []

    print(f"Processando {len(files)} arquivos com lógica v5 (Windows)...")
    for i, file_path in enumerate(files, start=1):
        if i == 1 or i % 50 == 0 or i == len(files):
            print(f"Progresso: {i}/{len(files)} - {file_path.name}")

        text = extract_text_from_pdf(file_path)
        data = parse_nf_data(text, file_path.name)
        results.append(data)

    df = pd.DataFrame(results)
    df.to_excel(OUTPUT_FILE, index=False)

    print("\nConferência concluída com sucesso!")
    print(f"Arquivo gerado em: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
