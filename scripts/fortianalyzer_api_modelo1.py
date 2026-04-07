


# Ajuste sys.path para garantir imports robustos
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import os
import requests
import json
import io
import base64
import zipfile
import urllib3
import pandas as pd
import re
from datetime import datetime, timedelta
from dotenv import load_dotenv
from services.recipients_service import RecipientsService

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Carrega variaveis do .env sem sobrescrever variaveis ja definidas no processo.
load_dotenv(override=False)
FAZ_IP = os.getenv("FAZ_IP", "10.254.12.34")
API_KEY = os.getenv("FAZ_API_KEY", "cnic6cm7yxp97nkh8wzku6cuspfc8t8g")
ADOM = os.getenv("FAZ_ADOM", "GPS_UNIDADES")
FAZ_API_URL = f"http://{FAZ_IP}/jsonrpc"

def faz_request(payload):
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {API_KEY}"
    }
    try:
        r = requests.post(
            FAZ_API_URL,
            headers=headers,
            json=payload,
            verify=False,
            timeout=30
        )
        return r.json()
    except Exception as e:
        print(f"Erro na requisição: {e}")
        return {}


def sanitize_filename(value):
    return str(value or "relatorio").replace("/", "_").replace(" ", "_")


def extract_result_name(resp_json):
    result = resp_json.get("result") if isinstance(resp_json, dict) else None
    if not isinstance(result, dict):
        return None

    value = result.get("name")
    if value is None:
        return None

    text = str(value).strip()
    return text or None


def extract_pdf_bytes_from_jsonrpc(resp_json):
    result = resp_json.get("result") if isinstance(resp_json, dict) else None
    if not isinstance(result, dict):
        return None, None, None

    raw_data = result.get("data")
    if not isinstance(raw_data, str) or len(raw_data) < 100:
        return None, None, None

    try:
        decoded_bytes = base64.b64decode(raw_data)
    except Exception:
        return None, None, None

    if decoded_bytes.startswith(b"%PDF"):
        return decoded_bytes, None, "pdf-base64"

    try:
        with zipfile.ZipFile(io.BytesIO(decoded_bytes)) as zip_file:
            members = zip_file.namelist()
            for member in members:
                if member.lower().endswith(".pdf"):
                    return zip_file.read(member), member, "zip-base64"
            return None, members, "zip-base64-sem-pdf"
    except zipfile.BadZipFile:
        return None, None, None


def extract_report_identity(value):
    text = str(value or "").strip()
    if text.lower().endswith(".pdf"):
        text = Path(text).stem
    match = re.match(r"^(.*?)-\d{4}-\d{2}-\d{2}-\d{4}-\d{4}(?:_\d+)?$", text)
    if match:
        text = match.group(1)
    text = text.upper()
    text = text.replace("-", " ").replace("_", " ")
    text = " ".join(text.split())
    return text


def matches_report_identity(candidate, expected):
    if not candidate or not expected:
        return False
    return candidate == expected or candidate.startswith(expected + " ")


def parse_report_period_end(report):
    value = report.get("period-end")
    if value:
        for fmt in ("%Y/%m/%d %H:%M", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(str(value).strip(), fmt)
            except ValueError:
                continue

    timestamp = report.get("timestamp-period-end")
    if timestamp is None:
        return None

    try:
        return datetime.fromtimestamp(int(timestamp))
    except Exception:
        return None


def report_covers_previous_month(report, target_last_day):
    period_end = parse_report_period_end(report)
    if period_end is None:
        return False
    return period_end.date() == target_last_day.date()


def main():
    # Calcula o primeiro e o último dia do mês ANTERIOR ao mês atual
    today = datetime.now()
    first_day_this_month = today.replace(day=1)
    last_day_prev_month = first_day_this_month - timedelta(days=1)
    first_day_prev_month = last_day_prev_month.replace(day=1)
    query_start = (first_day_this_month - timedelta(days=1)).strftime('%Y-%m-%d 00:00:00')
    query_end = today.strftime('%Y-%m-%d 23:59:59')


    # Caminhos dinâmicos para o mês/dia atual
    export_base = os.path.join("exports", today.strftime("%Y"), today.strftime("%b").lower(), today.strftime("%d"))
    pdf_dir = os.path.join(export_base, "pdf_do_mês")
    rascunhos_dir = os.path.join(export_base, "rascunhos_falha")
    os.makedirs(pdf_dir, exist_ok=True)
    os.makedirs(rascunhos_dir, exist_ok=True)

    # Carrega serviço de emails por regional
    lideres_path = os.path.join("data", "Lideres.xlsx")
    recipients_service = None
    if os.path.exists(lideres_path):
        try:
            recipients_service = RecipientsService(lideres_path, sheet_name="MODELOPY")
            print(f"Base de líderes carregada: {recipients_service.df.shape[0]} linhas")
        except Exception as e:
            print(f"Erro ao ler a planilha de líderes: {e}")
    else:
        print(f"Planilha de líderes não encontrada em {lideres_path}")

    print(
        f"\nListando relatórios disponíveis (/report/adom/{ADOM}/reports/state) "
        f"na janela de geração: {query_start} até {query_end} "
        f"para localizar PDFs com período encerrando em {last_day_prev_month:%Y-%m-%d}"
    )
    payload = {
        "id": 1,
        "jsonrpc": "2.0",
        "method": "get",
        "params": [
            {
                "apiver": 3,
                "adom": ADOM,
                "state": "generated",
                "url": f"/report/adom/{ADOM}/reports/state",
                "time-range": {
                    "start": query_start,
                    "end": query_end
                },
                "sort-by": [
                    { "field": "timestamp-start", "order": "desc" }
                ]
            }
        ]
    }
    resp = faz_request(payload)
    print(json.dumps(resp, indent=2, ensure_ascii=False))


    # Baixar PDF de cada relatório listado, filtrando apenas o lote cujo período se encerra no último dia do mês anterior.
    result = resp.get("result", []) if isinstance(resp, dict) else []
    if isinstance(result, list):
        reports = result[0].get("data", []) if result and isinstance(result[0], dict) else []
    elif isinstance(result, dict):
        reports = result.get("data", [])
    else:
        reports = []

    reports = [report for report in reports if report_covers_previous_month(report, last_day_prev_month)]
    pdfs_baixados = []
    for report in reports:
        tid = report.get("tid")
        title = report.get("title", "relatorio")
        report_name = report.get("name") or title
        # Normaliza o título para comparar de forma robusta
        def norm_title(s):
            import unicodedata
            s = str(s or "").strip().upper().replace("_", " ")
            s = unicodedata.normalize("NFKD", s)
            s = "".join(ch for ch in s if not unicodedata.combining(ch))
            s = " ".join(s.split())
            return s
        norm_tit = norm_title(title)
        # Aceita apenas "REPORT DE SEGURANCA" sem sufixo "-2" ou similares
        if not tid:
            continue
        if not (norm_tit == "REPORT DE SEGURANCA" or norm_tit == "REPORT DE SEGURANCA"):  # redundante para reforçar
            continue
        # Se tiver "-2" ou " 2" no final, pula
        if norm_tit.endswith("-2") or norm_tit.endswith(" 2"):
            continue
        print(f"\nBaixando PDF do relatório: {report_name} (tid: {tid})")
        download_payload = {
            "id": 2,
            "jsonrpc": "2.0",
            "method": "get",
            "params": [
                {
                    "apiver": 3,
                    "adom": ADOM,
                    "url": f"/report/adom/{ADOM}/reports/data/{tid}",
                    "format": "pdf"
                }
            ]
        }
        r = requests.post(
            FAZ_API_URL,
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {API_KEY}"
            },
            json=download_payload,
            verify=False,
            timeout=60
        )
        content_type = r.headers.get('Content-Type', '')
        # Se já veio PDF direto
        effective_report_name = report_name
        if r.status_code == 200 and r.content and 'application/pdf' in content_type:
            filename = f"{sanitize_filename(effective_report_name)}_{tid}.pdf"
            pdf_path = os.path.join(pdf_dir, filename)
            with open(pdf_path, "wb") as f:
                f.write(r.content)
            print(f"PDF salvo como: {pdf_path}")
            pdfs_baixados.append({"filename": filename, "path": pdf_path, "report_name": effective_report_name, "tid": tid})
        else:
            # Tenta decodificar como JSON para buscar o link/conteúdo do PDF
            try:
                resp_json = r.json()
            except Exception:
                resp_json = None
            pdf_downloaded = False
            if resp_json:
                result_name = extract_result_name(resp_json)
                if result_name:
                    effective_report_name = result_name
                pdf_bytes, pdf_member_name, payload_type = extract_pdf_bytes_from_jsonrpc(resp_json)
                if pdf_bytes:
                    filename_base = sanitize_filename(effective_report_name)
                    filename = f"{filename_base}_{tid}.pdf"
                    pdf_path = os.path.join(pdf_dir, filename)
                    with open(pdf_path, "wb") as f:
                        f.write(pdf_bytes)
                    origem = "ZIP interno do tid" if payload_type == "zip-base64" else "base64"
                    if pdf_member_name:
                        print(f"Arquivo interno encontrado no tid {tid}: {pdf_member_name}")
                    print(f"PDF salvo como: {pdf_path} (extraído de {origem})")
                    pdfs_baixados.append({
                        "filename": filename,
                        "path": pdf_path,
                        "report_name": effective_report_name,
                        "tid": tid,
                    })
                    pdf_downloaded = True

                # Tenta encontrar campo com link para download
                pdf_url = None
                for key in ['file', 'download_url', 'url', 'data']:
                    if key in resp_json:
                        pdf_url = resp_json[key]
                        break
                if not pdf_url:
                    for key in ['result', 'data', 'payload']:
                        if key in resp_json and isinstance(resp_json[key], dict):
                            for subkey in ['file', 'download_url', 'url']:
                                if subkey in resp_json[key]:
                                    pdf_url = resp_json[key][subkey]
                                    break
                        if pdf_url:
                            break
                if pdf_url and isinstance(pdf_url, str) and pdf_url.lower().endswith('.pdf') and not pdf_downloaded:
                    try:
                        pdf_resp = requests.get(pdf_url, headers={"Authorization": f"Bearer {API_KEY}"}, verify=False, timeout=60)
                        if pdf_resp.status_code == 200 and 'application/pdf' in pdf_resp.headers.get('Content-Type', ''):
                            filename = f"{sanitize_filename(effective_report_name)}_{tid}.pdf"
                            pdf_path = os.path.join(pdf_dir, filename)
                            with open(pdf_path, "wb") as f:
                                f.write(pdf_resp.content)
                            print(f"PDF salvo como: {pdf_path}")
                            pdfs_baixados.append({
                                "filename": filename,
                                "path": pdf_path,
                                "report_name": effective_report_name,
                                "tid": tid,
                            })
                            pdf_downloaded = True
                        else:
                            print(f"[ERRO] Falha ao baixar PDF do link encontrado: {pdf_url}")
                    except Exception as e:
                        print(f"[ERRO] Exceção ao baixar PDF do link: {e}")
                elif payload_type == "zip-base64-sem-pdf":
                    print(f"[ERRO] O tid {tid} retornou um ZIP sem PDF. Itens encontrados: {pdf_member_name}")
            if not pdf_downloaded:
                # Salva resposta para debug
                debug_filename = f"{sanitize_filename(report_name)}_{tid}_DEBUG.txt"
                debug_path = os.path.join(pdf_dir, debug_filename)
                with open(debug_path, "wb") as f:
                    f.write(r.content)
                print(f"[ERRO] Falha ao baixar PDF do relatório {report_name} (tid: {tid}) - Status: {r.status_code} - Content-Type: {content_type}")
                print(f"[ERRO] Resposta salva para debug em: {debug_path}")
                print(f"[ERRO] Início do conteúdo: {r.content[:200]}")


    # --- Função de normalização robusta ---
    import unicodedata
    def normaliza_nome(s):
        s = str(s or "").strip().upper()
        s = unicodedata.normalize("NFKD", s)
        s = "".join(ch for ch in s if not unicodedata.combining(ch))
        s = s.replace("-", " ").replace("_", " ")
        s = " ".join(s.split())
        return s

    # Extrai a "regional" do nome do PDF: tudo entre "REPORT DE SEGURANCA-" e o próximo hífen antes da data
    def extrai_regional_do_nome_pdf(nome_pdf):
        # Exemplo: REPORT DE SEGURANCA-FTG_ORMEC_PARA-2026-02-01-0127-0300_3050.pdf
        m = re.match(r"REPORT DE SEGURANCA-([^-]+)-", nome_pdf)
        if m:
            return normaliza_nome(m.group(1))
        return None

    # Pré-carrega todas as regionais normalizadas da planilha
    regionais_forti = {}
    if recipients_service:
        col_forti = recipients_service._find_col("NOME_REG_FORTI")
        col_regional = recipients_service._find_col("NOME_REGIONAL", "REGIONAL", "INTEGRADA", "REGIAO", "REGIÃO", "UF")
        if not col_forti:
            print("[ERRO] Coluna NOME_REG_FORTI não encontrada na planilha. Colunas disponíveis:", list(recipients_service.df.columns))
        else:
            for _, row in recipients_service.df.iterrows():
                forti_valor = row.get(col_forti)
                if pd.isna(forti_valor):
                    continue
                regional_valor = row.get(col_regional) if col_regional else forti_valor
                regionais_forti[normaliza_nome(forti_valor)] = {
                    "forti": str(forti_valor).strip(),
                    "regional": str(regional_valor).strip() if regional_valor is not None else str(forti_valor).strip(),
                }

    cruzamento = []
    if recipients_service and regionais_forti:
        for pdf in pdfs_baixados:
            source_name = pdf.get("report_name") or pdf["filename"]
            regional_pdf = extrai_regional_do_nome_pdf(source_name)
            emails = []
            regional_planilha = None
            regional_forti = None
            report_identity = extract_report_identity(source_name)
            regional_info = regionais_forti.get(report_identity)
            if not regional_info and report_identity:
                for forti_identity, candidate in regionais_forti.items():
                    if matches_report_identity(report_identity, forti_identity):
                        regional_info = candidate
                        break
            if regional_info:
                regional_planilha = regional_info.get("regional")
                regional_forti = regional_info.get("forti")
                emails = recipients_service.get_emails_by_regional(regional_planilha)
            elif regional_pdf:
                regional_info = regionais_forti.get(regional_pdf)
                if regional_info:
                    regional_planilha = regional_info.get("regional")
                    regional_forti = regional_info.get("forti")
                    emails = recipients_service.get_emails_by_regional(regional_planilha)
            cruzamento.append({
                "pdf": pdf["path"],
                "filename": pdf["filename"],
                "tid": pdf.get("tid"),
                "report_name": pdf.get("report_name"),
                "report_identity": report_identity,
                "regional_pdf": regional_pdf,
                "regional_planilha": regional_planilha,
                "regional_forti": regional_forti,
                "emails": emails
            })
        print("\nResumo do cruzamento PDF x Regional x Emails:")
        for item in cruzamento:
            print(f"PDF: {item['pdf']} | Regional PDF: {item['regional_pdf']} | Regional Planilha: {item['regional_planilha']} | Emails: {item['emails']}")
    else:
        print("Serviço de emails por regional não disponível ou coluna NOME_REG_FORTI ausente. Não foi possível cruzar PDFs com emails.")

    # Ativar modo teste se DRY_RUN ou SAFE_TEST_TO estiverem definidos
    dry_run = os.getenv('DRY_RUN', '1') == '1'
    safe_test_to = os.getenv('SAFE_TEST_TO')
    if dry_run:
        print("Modo DRY_RUN ativado: Nenhum email será enviado.")
    elif safe_test_to:
        print(f"Modo SAFE_TEST_TO ativado: Todos os emails serão enviados para {safe_test_to}.")
    else:
        print("Modo produção: Emails reais serão enviados.")

    return cruzamento

if __name__ == "__main__":
    main()