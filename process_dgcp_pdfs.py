#!/usr/bin/env python3
"""
Script to process DGCP PDF files, extract amortization payments,
and consolidate them into a formatted Excel file.
"""

import os
import re
import argparse
import glob
import pdfplumber
import pandas as pd

# ==============================================================================
# CONFIGURACIÓN DE RUTAS (Puedes editar estas variables directamente en el script)
# ==============================================================================
# Carpeta de entrada que contiene los archivos PDF (getjobid*.pdf)
INPUT_DIR = "."

# Nombre/ruta del archivo Excel de salida consolidado
OUTPUT_FILE = "pagos_amortizacion_consolidado.xlsx"
# ==============================================================================

def clean_value(val_str):
    """
    Cleans a numeric string by removing non-numeric characters (except minus, period, and comma),
    and converts it to float.
    """
    cleaned = re.sub(r'[^\d\.\,\-]', '', val_str)
    if not cleaned:
        return 0.0
    try:
        return float(cleaned.replace(',', ''))
    except ValueError:
        return 0.0

def parse_table_row(line):
    """
    Parses a single line from the flow table into the 8 expected columns.
    Expected: [Venc_Izq, Desembolsos, Amortizaciones, Intereses, Comisiones, Saldo_Pactado, Ref, Fecha_Der]
    """
    line_strip = line.strip()
    # Find all date strings (YYYY-MM-DD)
    dates = re.findall(r'\d{4}-\d{2}-\d{2}', line_strip)
    if not dates:
        return None

    venc_izq = dates[0]
    # Use the rightmost date as Fecha_Der
    fecha_der = dates[-1] if len(dates) >= 2 else venc_izq

    # Tokenize the line to extract clean values
    raw_tokens = line_strip.split()

    # Exclude Venc_Izq from content
    content_tokens = raw_tokens[1:]
    # If the last token is indeed the rightmost date, exclude it as well
    if len(dates) >= 2 and raw_tokens[-1] == fecha_der:
        content_tokens = content_tokens[:-1]

    # Filter out standalone non-numeric markers
    noise_patterns = [r'^P$', r'^D$', r'^\*+$', r'^Gu\u00eda-\d+$', r'^Guia-\d+$']
    filtered_tokens = []
    for tok in content_tokens:
        is_noise = False
        for pat in noise_patterns:
            if re.match(pat, tok, re.IGNORECASE):
                is_noise = True
                break
        if not is_noise:
            filtered_tokens.append(tok)

    numbers = []
    ref = ''

    for tok in filtered_tokens:
        # Clean potential attached markers from numbers
        clean_tok = re.sub(r'[P\*D]', '', tok).strip()
        if re.match(r'^-?[\d,]+(?:\.\d+)?$', clean_tok):
            numbers.append(clean_value(tok))
        else:
            ref = tok

    # Handle column padding/mapping
    if len(numbers) < 5:
        # Pad with 0.0 if not enough columns extracted
        numbers += [0.0] * (5 - len(numbers))
    elif len(numbers) == 6:
        # In case reference is a pure number and parsed as the 6th number
        ref = str(filtered_tokens[-1])
        numbers = numbers[:5]

    desembolsos, amortizaciones, intereses, comisiones, saldo_pactado = numbers[:5]

    return {
        'Venc_Izq': venc_izq,
        'Desembolsos': desembolsos,
        'Amortizaciones': amortizaciones,
        'Intereses': intereses,
        'Comisiones': comisiones,
        'Saldo_Pactado': saldo_pactado,
        'Ref': ref,
        'Fecha_Der': fecha_der
    }

def process_pdf_file(filepath):
    """
    Processes a single PDF file, extracts Credit, Tramos, and flow data,
    filters the rows with positive amortization, and returns them as a list of dicts.
    """
    records = []

    with pdfplumber.open(filepath) as pdf:
        credito = None
        tramos_registry = {} # maps tramo_id to (monto_total, currency)

        # Pass 1: Build the tramos_registry and find Crédito
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            lines = text.split('\n')
            for i, l in enumerate(lines):
                l_strip = l.strip()

                # Extract Crédito (e.g. "Crédito: 640100010" or "HOJA DE VIDA - CREDITO 512100103")
                cred_match = re.search(r'CREDITO\s+(\d+)', l_strip, re.IGNORECASE) or re.search(r'Cr\u00e9dito:\s*(\d+)', l_strip, re.IGNORECASE)
                if cred_match:
                    credito = cred_match.group(1)

                # Extract Tramo, Monto Total of Tramo, and Currency
                # Style A: "Tramo 178 33,667,171,200,000.00 - COP - TFIT15260826 - TES B"
                m_same = re.search(r'Tramo\s+(\d+)\s+([\d,]+(?:\.\d+)?)\s*-\s*([A-Z]{3})', l_strip, re.IGNORECASE)
                # Style B: "Tramo ( 1 )" with amount on the next line
                m_paren = re.search(r'Tramo\s*\(\s*(\d+)\s*\)', l_strip, re.IGNORECASE)

                if m_same:
                    tid = int(m_same.group(1))
                    amt = float(m_same.group(2).replace(',', ''))
                    curr = m_same.group(3)
                    tramos_registry[tid] = (amt, curr)
                elif m_paren:
                    tid = int(m_paren.group(1))
                    if i + 1 < len(lines):
                        next_line = lines[i+1].strip()
                        m_next = re.match(r'^([\d,]+(?:\.\d+)?)\s*-\s*([A-Z]{3})', next_line)
                        if m_next:
                            amt = float(m_next.group(1).replace(',', ''))
                            curr = m_next.group(2)
                            tramos_registry[tid] = (amt, curr)

        if not credito:
            # Fallback if no Credit ID found
            filename = os.path.basename(filepath)
            print(f"Warning: Crédito ID not found in {filename}")
            credito = "Unknown"

        # Pass 2: Parse flows
        current_tramo_id = None
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            lines = text.split('\n')
            for i, l in enumerate(lines):
                l_strip = l.strip()

                # Identify current active tramo context
                m_same = re.search(r'Tramo\s+(\d+)\s+([\d,]+(?:\.\d+)?)\s*-\s*([A-Z]{3})', l_strip, re.IGNORECASE)
                m_paren = re.search(r'Tramo\s*\(\s*(\d+)\s*\)', l_strip, re.IGNORECASE)
                m_simple = re.search(r'^Tramo\s+(\d+)$', l_strip, re.IGNORECASE)

                if m_same:
                    current_tramo_id = int(m_same.group(1))
                elif m_paren:
                    current_tramo_id = int(m_paren.group(1))
                elif m_simple:
                    current_tramo_id = int(m_simple.group(1))

                # Identify and parse flow table row
                if re.match(r'^\d{4}-\d{2}-\d{2}', l_strip) and ('Total:' not in l_strip) and ('VENCIMIENTO' not in l_strip) and ('HOJA' not in l_strip):
                    # Filter out top-of-page dates which have no flow details
                    dates = re.findall(r'\d{4}-\d{2}-\d{2}', l_strip)
                    if len(dates) == 1 and len(l_strip.split()) == 1:
                        continue

                    parsed = parse_table_row(l_strip)
                    if parsed:
                        monto_amortizado = parsed['Amortizaciones']
                        # Filtro estricto: keep only rows where Amortizaciones > 0
                        if monto_amortizado > 0:
                            tid = current_tramo_id if current_tramo_id is not None else 1

                            # Look up corresponding tramo header details
                            monto_total_tramo = 0.0
                            if tid in tramos_registry:
                                monto_total_tramo = tramos_registry[tid][0]
                            else:
                                if len(tramos_registry) == 1:
                                    monto_total_tramo = list(tramos_registry.values())[0][0]

                            records.append({
                                'Credito': int(credito) if credito.isdigit() else credito,
                                'Tramo': tid,
                                'Monto_Total_Tramo': monto_total_tramo,
                                'Fecha_Amortizacion': parsed['Fecha_Der'],
                                'Monto_Amortizado': monto_amortizado,
                                'Documento_Ref': parsed['Ref']
                            })

    return records

def main():
    parser = argparse.ArgumentParser(description="Process DGCP PDFs and consolidate amortization payments to Excel.")
    parser.add_argument("--input-dir", default=INPUT_DIR, help=f"Directory containing getjobid*.pdf files (default: {INPUT_DIR})")
    parser.add_argument("--output", default=OUTPUT_FILE, help=f"Path for the output Excel file (default: {OUTPUT_FILE})")

    args = parser.parse_args()
    print(f"Input directory: {args.input_dir}")
    print(f"Output path: {args.output}")

    # Find all PDFs
    pdf_pattern = os.path.join(args.input_dir, "getjobid*.pdf")
    pdf_files = sorted(glob.glob(pdf_pattern))

    if not pdf_files:
        print(f"No files matching 'getjobid*.pdf' found in '{args.input_dir}'.")
        return

    print(f"Found {len(pdf_files)} PDF file(s) to process.")

    all_records = []
    for pdf_file in pdf_files:
        print(f"Processing: {os.path.basename(pdf_file)}...")
        try:
            records = process_pdf_file(pdf_file)
            print(f"  -> Extracted {len(records)} amortization payments.")
            all_records.extend(records)
        except Exception as e:
            print(f"  -> Error processing {pdf_file}: {e}")

    print(f"Total extracted payments (including duplicates): {len(all_records)}")

    if not all_records:
        print("No payments found to write to Excel.")
        return

    # Convert to DataFrame
    df = pd.DataFrame(all_records)

    # Remove duplicates (e.g. if the same PDF exists twice in the folder)
    df = df.drop_duplicates()
    print(f"Total unique payments after deduplication: {len(df)}")

    # Option B Calculation:
    # Calculate Porcentaje_Amortizado grouping by Credito and Tramo over the final consolidated dataset
    df['Porcentaje_Amortizado'] = df.groupby(['Credito', 'Tramo'])['Monto_Amortizado'].transform(
        lambda x: x / x.sum() if x.sum() > 0 else 0.0
    )

    # Order columns as specified
    cols = ['Credito', 'Tramo', 'Monto_Total_Tramo', 'Fecha_Amortizacion', 'Monto_Amortizado', 'Porcentaje_Amortizado', 'Documento_Ref']
    df = df[cols]

    # Save to Excel
    print(f"Saving to Excel: {args.output}...")
    df.to_excel(args.output, index=False)
    print("Done!")

if __name__ == "__main__":
    main()
