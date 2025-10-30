# confronto_de_saldos_web_v18.py
# Streamlit web app com OCR automático + fallback AgGrid + export Excel/PDF
# Compatível com Python 3.13 (Streamlit Cloud)
# Uso local: streamlit run confronto_de_saldos_web_v18.py

import io
import re
import pandas as pd
import streamlit as st
from typing import List, Tuple

# ===============================
# 1️⃣ AgGrid (fallback automático)
# ===============================
try:
    from st_aggrid import AgGrid, GridOptionsBuilder, JsCode
    from st_aggrid.shared import GridUpdateMode
    AGGRID_AVAILABLE = True
except ImportError:
    AGGRID_AVAILABLE = False

# ===============================
# 2️⃣ OCR / PDF / EXPORT
# ===============================
try:
    import pdfplumber
except Exception:
    pdfplumber = None

try:
    from pdf2image import convert_from_bytes
except Exception:
    convert_from_bytes = None

try:
    import pytesseract
except Exception:
    pytesseract = None

try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
except Exception:
    pass

# ===============================
# 3️⃣ EXPRESSÕES / PADRÕES
# ===============================
CURRENCY_RE = re.compile(r'(-?\d{1,3}(?:[.\s]\d{3})*(?:,\d{2})|\d+(?:,\d{2}))')
LABELS_PREV = [r'SALDO ANTERIOR', r'Saldo anterior', r'Saldo Anterior']
LABELS_CURR = [r'SALDO ATUAL', r'Saldo atual', r'Saldo Atual']

# ===============================
# 4️⃣ FUNÇÕES UTILITÁRIAS
# ===============================
def parse_currency_to_float(s: str):
    if s is None:
        return None
    s = str(s)
    s = re.sub(r'[Rr]\$\s*', '', s).strip()
    m = CURRENCY_RE.search(s)
    if not m:
        s2 = re.sub(r'[^\d\.\-]', '', s)
        try:
            return float(s2) if s2 else None
        except:
            return None
    token = m.group(0)
    token = token.replace('.', '').replace(' ', '').replace(',', '.')
    try:
        return float(token)
    except:
        return None


def format_brazilian(value):
    if pd.isna(value):
        return ''
    try:
        if isinstance(value, (int, float)):
            return f"{value:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
        else:
            v = float(str(value).replace('.', '').replace(',', '.'))
            return f"{v:,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
    except Exception:
        return str(value)


def extract_text_with_pdfplumber_bytes(b: bytes) -> str:
    if not pdfplumber:
        return ""
    try:
        with pdfplumber.open(io.BytesIO(b)) as pdf:
            pages = [p.extract_text() or '' for p in pdf.pages]
            return "\n".join(pages)
    except Exception:
        return ""


def extract_text_with_ocr_bytes(b: bytes) -> str:
    if convert_from_bytes is None or pytesseract is None:
        return ""
    try:
        images = convert_from_bytes(b)
    except Exception:
        return ""
    texts = []
    for img in images:
        try:
            txt = pytesseract.image_to_string(img, lang='por+eng')
        except Exception:
            txt = ""
        texts.append(txt)
    return "\n".join(texts)


def extract_text_from_pdf_bytes_smart(b: bytes) -> str:
    text = extract_text_with_pdfplumber_bytes(b)
    if text and len(text.strip()) > 20:
        return text
    if pytesseract and convert_from_bytes:
        return extract_text_with_ocr_bytes(b)
    return text or ""


def find_label_value_in_window(window: str, labels):
    for lab in labels:
        m = re.search(r'(' + re.escape(lab) + r')[^\n\r]{0,100}([^\n\r]*)', window)
        if m:
            tail = m.group(2)
            val = parse_currency_to_float(tail)
            if val is not None:
                return val
    return None


def find_balance_by_label_in_text(text: str, account: str, label_group='current'):
    if not text:
        return None
    txt = text
    acct_esc = re.escape(account.strip())
    labels = LABELS_CURR if label_group == 'current' else LABELS_PREV

    for m in re.finditer(acct_esc, txt):
        start = max(0, m.start() - 800)
        end = min(len(txt), m.end() + 800)
        window = txt[start:end]
        val = find_label_value_in_window(window, labels)
        if val is not None:
            return val

    val_global = find_label_value_in_window(txt, labels)
    if val_global is not None:
        return val_global

    for lab in labels:
        m2 = re.search(re.escape(lab) + r'\s*[=:]?\s*([Rr]?\$?[^\n\r]{0,40})', txt)
        if m2:
            val = parse_currency_to_float(m2.group(1))
            if val is not None:
                return val
    return None

# ===============================
# 5️⃣ PROCESSAMENTO PRINCIPAL
# ===============================
def process_confronto_streamlit(df_accounts: pd.DataFrame, pdf_files_bytes: List[bytes]):
    results = []
    total_prev_excel = total_prev_pdf = total_curr_excel = total_curr_pdf = 0.0
    total_accounts = len(df_accounts.index)
    processed = 0

    for idx, row in enumerate(df_accounts.itertuples(index=False), start=1):
        if st.session_state.get('cancel_requested', False):
            return results, (total_prev_excel, total_prev_pdf, total_curr_excel, total_curr_pdf), True

        account = str(getattr(row, 'account')).strip()
        excel_prev = getattr(row, 'excel_prev')
        excel_curr = getattr(row, 'excel_curr')

        if excel_prev: total_prev_excel += excel_prev
        if excel_curr: total_curr_excel += excel_curr

        pdf_prev = pdf_curr = None
        for b in pdf_files_bytes:
            if st.session_state.get('cancel_requested', False):
                return results, (total_prev_excel, total_prev_pdf, total_curr_excel, total_curr_pdf), True
            text = extract_text_from_pdf_bytes_smart(b)
            if not text: continue
            if account in text:
                pdf_prev = find_balance_by_label_in_text(text, account, 'previous')
                pdf_curr = find_balance_by_label_in_text(text, account, 'current')
                if pdf_prev is not None and pdf_curr is not None:
                    break

        if pdf_prev: total_prev_pdf += pdf_prev
        if pdf_curr: total_curr_pdf += pdf_curr

        def cmp(a, b):
            if a is None or b is None: return None, 'unk'
            diff = round(a - b, 2)
            return diff, 'confere' if abs(diff) <= 0.01 else 'nao_confere'

        diff_prev, tag_prev = cmp(excel_prev, pdf_prev)
        diff_curr, tag_curr = cmp(excel_curr, pdf_curr)

        res = dict(
            account=account,
            excel_prev=excel_prev,
            pdf_prev=pdf_prev,
            diff_prev=diff_prev,
            status_prev='CONFERE' if tag_prev == 'confere' else ('NÃO CONFERE' if tag_prev == 'nao_confere' else 'NÃO ENCONTRADO'),
            excel_curr=excel_curr,
            pdf_curr=pdf_curr,
            diff_curr=diff_curr,
            status_curr='CONFERE' if tag_curr == 'confere' else ('NÃO CONFERE' if tag_curr == 'nao_confere' else 'NÃO ENCONTRADO'),
        )
        results.append(res)

        processed += 1
        percent = int((processed / total_accounts) * 100)
        st.session_state['progress_percent'] = percent
        st.session_state['processed_count'] = processed

    st.session_state['progress_percent'] = 100
    st.session_state['processed_count'] = processed
    return results, (total_prev_excel, total_prev_pdf, total_curr_excel, total_curr_pdf), False

# ===============================
# 6️⃣ EXPORTAÇÕES
# ===============================
def create_excel_bytes(results: List[dict], totals: Tuple[float, float, float, float]) -> bytes:
    df = pd.DataFrame(results)
    headers = {
        "account": "Conta", "excel_prev": "Excel Prev", "pdf_prev": "PDF Prev",
        "diff_prev": "Dif Prev", "status_prev": "Status Prev",
        "excel_curr": "Excel Curr", "pdf_curr": "PDF Curr",
        "diff_curr": "Dif Curr", "status_curr": "Status Curr"
    }
    df.rename(columns=headers, inplace=True)
    for col in ["Excel Prev", "PDF Prev", "Dif Prev", "Excel Curr", "PDF Curr", "Dif Curr"]:
        if col in df.columns: df[col] = df[col].apply(format_brazilian)

    total_prev_excel, total_prev_pdf, total_curr_excel, total_curr_pdf = totals
    df.loc[len(df)] = ["Totais", format_brazilian(total_prev_excel), format_brazilian(total_prev_pdf),
                       "", "", format_brazilian(total_curr_excel), format_brazilian(total_curr_pdf),
                       "", "", ]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Confronto de Saldos")
    output.seek(0)
    return output.read()


def create_pdf_bytes(results: List[dict], totals: Tuple[float, float, float, float]) -> bytes:
    header = ["Conta", "Excel Prev", "PDF Prev", "Dif Prev", "Status Prev",
              "Excel Curr", "PDF Curr", "Dif Curr", "Status Curr"]
    data = [header]
    styles = [('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#d3d3d3')),
              ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
              ('ALIGN', (0, 0), (-1, -1), 'CENTER')]

    for r in results:
        row = [r['account'], format_brazilian(r['excel_prev']), format_brazilian(r['pdf_prev']),
               format_brazilian(r['diff_prev']), r['status_prev'], format_brazilian(r['excel_curr']),
               format_brazilian(r['pdf_curr']), format_brazilian(r['diff_curr']), r['status_curr']]
        data.append(row)
        if "NÃO CONFERE" in r['status_prev'] or "NÃO CONFERE" in r['status_curr']:
            styles.append(('BACKGROUND', (0, len(data) - 1), (-1, len(data) - 1), colors.HexColor('#fce4e4')))

    total_prev_excel, total_prev_pdf, total_curr_excel, total_curr_pdf = totals
    totals_row = ["Totais", format_brazilian(total_prev_excel), format_brazilian(total_prev_pdf),
                  "", "", format_brazilian(total_curr_excel), format_brazilian(total_curr_pdf),
                  "", "", ]
    data.append(totals_row)
    styles.append(('BACKGROUND', (0, len(data)-1), (-1, len(data)-1), colors.HexColor('#EDEDED')))

    output = io.BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(A4))
    story = [Paragraph("Relatório de Confronto de Saldos", getSampleStyleSheet()['Title']), Spacer(1, 12),
             Table(data, style=TableStyle(styles), repeatRows=1)]
    doc.build(story)
    output.seek(0)
    return output.read()

# ===============================
# 7️⃣ INTERFACE STREAMLIT
# ===============================
st.set_page_config(page_title="Confronto de Saldos", layout="wide")
st.title("Confronto de Saldos — v18 (Web OCR + Fallback AgGrid)")

for k, v in [('progress_percent', 0), ('processed_count', 0), ('cancel_requested', False),
             ('last_results', []), ('last_totals', (0.0, 0.0, 0.0, 0.0))]:
    if k not in st.session_state: st.session_state[k] = v

with st.sidebar:
    st.header("Entradas")
    excel_file = st.file_uploader("Upload Excel (Conta | Saldo Anterior | Saldo Atual)", type=['xlsx'])
    pdf_files = st.file_uploader("Upload PDFs", type=['pdf'], accept_multiple_files=True)

    st.markdown("---")
    use_ocr = st.checkbox("Ativar OCR automático", value=True)
    st.info("Planilha: 1ª Conta, 2ª Saldo Anterior, 3ª Saldo Atual")

df_accounts = None
if excel_file:
    df = pd.read_excel(excel_file)
    df.columns = ['account', 'excel_prev', 'excel_curr']
    for col in ['excel_prev', 'excel_curr']:
        df[col] = df[col].apply(lambda x: parse_currency_to_float(x))
    df_accounts = df
    st.sidebar.success(f"Planilha carregada: {len(df)} contas.")

pdf_bytes_list = [f.read() for f in pdf_files] if pdf_files else []

col1, col2 = st.columns([1, 1])
with col1:
    run = st.button("Executar Confronto")
    cancel = st.button("Cancelar")
    if cancel:
        st.session_state['cancel_requested'] = True
with col2:
    st.progress(st.session_state['progress_percent'])
    st.write(f"{st.session_state['progress_percent']}% — {st.session_state['processed_count']} processadas")

if run:
    st.session_state.update(progress_percent=0, processed_count=0, cancel_requested=False)
    with st.spinner("Processando..."):
        results, totals, cancelled = process_confronto_streamlit(df_accounts, pdf_bytes_list)
        st.session_state['last_results'], st.session_state['last_totals'] = results, totals
        st.success("Confronto concluído!" if not cancelled else "Processo cancelado.")

results_display = st.session_state['last_results']
totals = st.session_state['last_totals']
st.markdown("---")
st.subheader("Resultados")

if results_display:
    df_display = pd.DataFrame(results_display)
    df_display_fmt = df_display.copy()
    for c in ['excel_prev', 'pdf_prev', 'diff_prev', 'excel_curr', 'pdf_curr', 'diff_curr']:
        df_display_fmt[c] = df_display_fmt[c].apply(format_brazilian)

    if AGGRID_AVAILABLE:
        gb = GridOptionsBuilder.from_dataframe(df_display_fmt)
        js_style = JsCode("""
        function(params){if(params.data['status_prev']=='NÃO CONFERE'||params.data['status_curr']=='NÃO CONFERE')
        {return {'backgroundColor':'#f7d4d4'};}return null;}""")
        gb.configure_default_column(cellStyle=js_style)
        AgGrid(df_display_fmt, gridOptions=gb.build(), height=400)
    else:
        st.warning("⚠️ Tabela interativa (AgGrid) indisponível — exibindo tabela padrão.")
        st.dataframe(df_display_fmt, use_container_width=True)

    st.markdown("**Totais (por fonte)**")
    c1, c2 = st.columns(2)
    with c1:
        st.write(f"Saldo Anterior (Excel): **{format_brazilian(totals[0])}**")
        st.write(f"Saldo Atual (Excel): **{format_brazilian(totals[2])}**")
    with c2:
        st.write(f"Saldo Anterior (PDF): **{format_brazilian(totals[1])}**")
        st.write(f"Saldo Atual (PDF): **{format_brazilian(totals[3])}**")

    st.markdown("---")
    colx, coly = st.columns(2)
    with colx:
        st.download_button("Exportar Excel", create_excel_bytes(results_display, totals),
                           "Relatorio_Confronto.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with coly:
        st.download_button("Exportar PDF", create_pdf_bytes(results_display, totals),
                           "Relatorio_Confronto.pdf", mime="application/pdf")
else:
    st.info("Faça upload do Excel e PDFs e clique em 'Executar Confronto'.")
