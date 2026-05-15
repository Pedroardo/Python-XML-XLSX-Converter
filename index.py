"""
api/index.py — Vercel Serverless Function
All routes are handled here via the Flask app.
Vercel calls this file automatically for every request.
"""

from flask import Flask, request, send_file, jsonify, render_template_string
import io
from datetime import datetime
from xml.etree import ElementTree as ET
from openpyxl import load_workbook

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024


# ── Conversion logic ──────────────────────────────────────────────────────────

JENIS_FAKTUR_MAP = {
    "Normal": "Normal", "normal": "Normal",
    "Pengganti": "Pengganti", "pengganti": "Pengganti",
}
COUNTRY_FIX = {"IND": "IDN"}

def fix_country(code): return COUNTRY_FIX.get(code, code)

def format_date(raw):
    try:
        return datetime.strptime(raw.strip(), "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return raw

def _text(el, tag, default=""):
    child = el.find(tag)
    if child is None or child.text is None:
        return default
    return child.text.strip()

def _num(el, tag, default="0"):
    val = _text(el, tag, default)
    try:
        f = float(val)
        return str(int(f)) if f == int(f) else str(f)
    except ValueError:
        return val

def parse_xml_bytes(xml_bytes):
    root = ET.fromstring(xml_bytes)
    seller_tin = _text(root, "TIN")
    invoices = []
    for inv in root.iter("TaxInvoice"):
        invoice = {
            "seller_tin":       seller_tin,
            "date":             format_date(_text(inv, "TaxInvoiceDate")),
            "jenis_faktur":     JENIS_FAKTUR_MAP.get(_text(inv, "TaxInvoiceOpt"), _text(inv, "TaxInvoiceOpt")),
            "kode_transaksi":   _text(inv, "TrxCode"),
            "ket_tambahan":     _text(inv, "AddInfo"),
            "dok_pendukung":    _text(inv, "CustomDoc"),
            "referensi":        _text(inv, "RefDesc"),
            "cap_fasilitas":    _text(inv, "FacilityStamp"),
            "id_tku_penjual":   _text(inv, "SellerIDTKU"),
            "npwp_pembeli":     _text(inv, "BuyerTin"),
            "jenis_id_pembeli": _text(inv, "BuyerDocument"),
            "negara_pembeli":   fix_country(_text(inv, "BuyerCountry")),
            "no_dok_pembeli":   _text(inv, "BuyerDocumentNumber") or "-",
            "nama_pembeli":     _text(inv, "BuyerName"),
            "alamat_pembeli":   _text(inv, "BuyerAdress"),
            "email_pembeli":    _text(inv, "BuyerEmail"),
            "id_tku_pembeli":   _text(inv, "BuyerIDTKU"),
        }
        goods = []
        for gs in inv.iter("GoodService"):
            goods.append({
                "opt":         _text(gs, "Opt"),
                "kode":        _text(gs, "Code"),
                "nama":        _text(gs, "Name"),
                "satuan":      _text(gs, "Unit"),
                "harga":       _num(gs, "Price"),
                "qty":         _num(gs, "Qty"),
                "diskon":      _num(gs, "TotalDiscount"),
                "dpp":         _num(gs, "TaxBase"),
                "dpp_lain":    _num(gs, "OtherTaxBase"),
                "tarif_ppn":   _num(gs, "VATRate"),
                "ppn":         _num(gs, "VAT"),
                "tarif_ppnbm": _num(gs, "STLGRate"),
                "ppnbm":       _num(gs, "STLG"),
            })
        invoices.append((invoice, goods))
    return invoices

def write_xlsx_bytes(invoices, template_bytes):
    buf = io.BytesIO(template_bytes)
    wb  = load_workbook(buf)
    ws_faktur = wb["Faktur"]
    ws_detail = wb["DetailFaktur"]

    def find_header_row(ws):
        for r in ws.iter_rows():
            for cell in r:
                if cell.value == "Baris":
                    return cell.row
        return None

    fhr = find_header_row(ws_faktur)
    dhr = find_header_row(ws_detail)

    def clear_data_rows(ws, start):
        rows_to_delete = []
        for row in ws.iter_rows(min_row=start):
            if any(c.value is not None for c in row):
                rows_to_delete.append(row[0].row)
        for r in reversed(rows_to_delete):
            ws.delete_rows(r)

    clear_data_rows(ws_faktur, fhr + 1)
    clear_data_rows(ws_detail, dhr + 1)

    def col_map(ws, hr):
        return {str(c.value).strip(): c.column for c in ws[hr] if c.value}

    f_cols = col_map(ws_faktur, fhr)
    d_cols = col_map(ws_detail, dhr)

    if invoices:
        for row in ws_faktur.iter_rows(max_row=fhr - 1):
            for cell in row:
                if cell.value and "NPWP" in str(cell.value).upper() and "PENJUAL" in str(cell.value).upper():
                    for c in range(cell.column + 1, cell.column + 10):
                        t = ws_faktur.cell(row=cell.row, column=c)
                        if type(t).__name__ == "Cell":
                            t.value = invoices[0][0]["seller_tin"]
                            break
                    break

    def sc(ws, row, name, val, cm):
        col = cm.get(name)
        if col:
            ws.cell(row=row, column=col).value = val

    f_row = fhr + 1
    d_row = dhr + 1
    for baris, (inv, goods) in enumerate(invoices, 1):
        sc(ws_faktur, f_row, "Baris",                 baris,                   f_cols)
        sc(ws_faktur, f_row, "Tanggal Faktur",        inv["date"],              f_cols)
        sc(ws_faktur, f_row, "Jenis Faktur",          inv["jenis_faktur"],      f_cols)
        sc(ws_faktur, f_row, "Kode Transaksi",        inv["kode_transaksi"],    f_cols)
        sc(ws_faktur, f_row, "Keterangan Tambahan",   inv["ket_tambahan"],      f_cols)
        sc(ws_faktur, f_row, "Dokumen Pendukung",     inv["dok_pendukung"],     f_cols)
        sc(ws_faktur, f_row, "Referensi",             inv["referensi"],         f_cols)
        sc(ws_faktur, f_row, "Cap Fasilitas",         inv["cap_fasilitas"],     f_cols)
        sc(ws_faktur, f_row, "ID TKU Penjual",        inv["id_tku_penjual"],    f_cols)
        sc(ws_faktur, f_row, "NPWP/NIK Pembeli",     inv["npwp_pembeli"],      f_cols)
        sc(ws_faktur, f_row, "Jenis ID Pembeli",      inv["jenis_id_pembeli"],  f_cols)
        sc(ws_faktur, f_row, "Negara Pembeli",        inv["negara_pembeli"],    f_cols)
        sc(ws_faktur, f_row, "Nomor Dokumen Pembeli", inv["no_dok_pembeli"],    f_cols)
        sc(ws_faktur, f_row, "Nama Pembeli",          inv["nama_pembeli"],      f_cols)
        sc(ws_faktur, f_row, "Alamat Pembeli",        inv["alamat_pembeli"],    f_cols)
        sc(ws_faktur, f_row, "Email Pembeli",         inv["email_pembeli"],     f_cols)
        sc(ws_faktur, f_row, "ID TKU Pembeli",        inv["id_tku_pembeli"],    f_cols)
        f_row += 1
        for gs in goods:
            sc(ws_detail, d_row, "Baris",              baris,             d_cols)
            sc(ws_detail, d_row, "Barang/Jasa",        gs["opt"],         d_cols)
            sc(ws_detail, d_row, "Kode Barang Jasa",   gs["kode"],        d_cols)
            sc(ws_detail, d_row, "Nama Barang/Jasa",   gs["nama"],        d_cols)
            sc(ws_detail, d_row, "Nama Satuan Ukur",   gs["satuan"],      d_cols)
            sc(ws_detail, d_row, "Harga Satuan",       gs["harga"],       d_cols)
            sc(ws_detail, d_row, "Jumlah Barang Jasa", gs["qty"],         d_cols)
            sc(ws_detail, d_row, "Total Diskon",       gs["diskon"],      d_cols)
            sc(ws_detail, d_row, "DPP",                gs["dpp"],         d_cols)
            sc(ws_detail, d_row, "DPP Nilai Lain",     gs["dpp_lain"],    d_cols)
            sc(ws_detail, d_row, "Tarif PPN",          gs["tarif_ppn"],   d_cols)
            sc(ws_detail, d_row, "PPN",                gs["ppn"],         d_cols)
            sc(ws_detail, d_row, "Tarif PPnBM",        gs["tarif_ppnbm"], d_cols)
            sc(ws_detail, d_row, "PPnBM",              gs["ppnbm"],       d_cols)
            d_row += 1

    ws_faktur.cell(row=f_row, column=f_cols.get("Baris", 1)).value = "END"
    ws_detail.cell(row=d_row, column=d_cols.get("Baris", 1)).value = "END"

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ── HTML ──────────────────────────────────────────────────────────────────────

HTML = """<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Coretax → PajakExpress Converter</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');
  :root {
    --bg: #080b10; --card: #0e1318; --border: #1e2830;
    --accent: #00d4ff; --green: #00ff9d;
    --text: #c8d8e8; --sub: #4a6070; --danger: #ff4d6d;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    background: var(--bg); color: var(--text);
    font-family: 'IBM Plex Sans', sans-serif;
    min-height: 100vh; display: flex; flex-direction: column;
    align-items: center; padding: 40px 16px 60px;
  }
  body::before {
    content: ''; position: fixed; inset: 0; pointer-events: none; z-index: 0;
    background-image: linear-gradient(rgba(0,212,255,.04) 1px, transparent 1px),
                      linear-gradient(90deg, rgba(0,212,255,.04) 1px, transparent 1px);
    background-size: 40px 40px;
  }
  .wrap { position: relative; z-index: 1; width: 100%; max-width: 560px; }
  header { text-align: center; margin-bottom: 36px; }
  .logo { font-size: 13px; color: var(--accent); letter-spacing: 4px;
          font-family: 'IBM Plex Mono', monospace; margin-bottom: 12px; }
  h1 { font-size: 26px; font-weight: 600; color: #fff; line-height: 1.2; }
  h1 span { color: var(--accent); }
  .sub-h { font-size: 13px; color: var(--sub); margin-top: 6px; }
  .card { background: var(--card); border: 1px solid var(--border);
          border-radius: 6px; padding: 24px; margin-bottom: 16px; }
  .section-label { font-family: 'IBM Plex Mono', monospace; font-size: 10px;
                   color: var(--sub); letter-spacing: 2px; margin-bottom: 16px; }
  .field + .field { margin-top: 14px; }
  label { display: block; font-size: 12px; color: var(--sub); margin-bottom: 6px; }
  .file-row { display: flex; gap: 8px; }
  .file-input {
    flex: 1; background: #050709; border: 1px solid var(--border);
    border-radius: 4px; color: var(--text);
    font-family: 'IBM Plex Mono', monospace; font-size: 11px;
    padding: 9px 12px; outline: none; cursor: pointer; transition: border-color .2s;
  }
  .file-input:hover { border-color: var(--accent); }
  .stats { display: flex; gap: 12px; margin-bottom: 14px; }
  .stat { flex: 1; background: #050709; border: 1px solid var(--border);
          border-radius: 4px; padding: 12px; text-align: center; }
  .stat-num { font-size: 28px; font-weight: 600; color: var(--accent);
              font-family: 'IBM Plex Mono', monospace; }
  .stat-num.green { color: var(--green); }
  .stat-lbl { font-size: 10px; color: var(--sub); margin-top: 2px; }
  .log { background: #020304; border: 1px solid var(--border); border-radius: 4px;
         font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: var(--green);
         padding: 12px; height: 120px; overflow-y: auto; white-space: pre-wrap;
         word-break: break-all; }
  .log .err { color: var(--danger); }
  .convert-btn {
    width: 100%; background: var(--accent); color: #000; border: none;
    border-radius: 4px; font-family: 'IBM Plex Mono', monospace;
    font-size: 13px; font-weight: 600; letter-spacing: 2px;
    padding: 16px; cursor: pointer; transition: opacity .2s, transform .1s;
  }
  .convert-btn:hover  { opacity: .88; }
  .convert-btn:active { transform: scale(.99); }
  .convert-btn:disabled { opacity: .4; cursor: not-allowed; }
  .success-banner {
    display: none; background: rgba(0,255,157,.08);
    border: 1px solid var(--green); border-radius: 4px;
    padding: 14px 18px; color: var(--green); font-size: 13px;
    margin-bottom: 16px; text-align: center;
  }
  footer { margin-top: 40px; font-size: 11px; color: var(--sub); text-align: center; }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div class="logo">PAJAK CONVERTER</div>
    <h1>XML <span>→</span> XLSX</h1>
    <p class="sub-h">Konversi XML / CSV Coretax ke template XLSX PajakExpress</p>
  </header>

  <div id="success-banner" class="success-banner"></div>

  <div class="card">
    <div class="section-label">01 / INPUT FILES</div>
    <div class="field">
      <label>File XML *</label>
      <div class="file-row">
        <input class="file-input" type="file" id="xml-file" accept=".xml" onchange="onXmlChange(this)">
      </div>
    </div>
    <div class="field">
      <label>Template XLSX PajakExpress *</label>
      <div class="file-row">
        <input class="file-input" type="file" id="tmpl-file" accept=".xlsx">
      </div>
    </div>
  </div>

  <div class="card">
    <div class="section-label">02 / PREVIEW</div>
    <div class="stats">
      <div class="stat">
        <div class="stat-num" id="inv-count">—</div>
        <div class="stat-lbl">Jumlah Faktur</div>
      </div>
      <div class="stat">
        <div class="stat-num green" id="item-count">—</div>
        <div class="stat-lbl">Jumlah Item</div>
      </div>
    </div>
    <div class="log" id="log">Pilih file XML untuk melihat preview…</div>
  </div>

  <button class="convert-btn" id="conv-btn" onclick="doConvert()">CONVERT</button>
  <footer>Coretax DLP XML → PajakExpress XLSX &nbsp;|</footer>
</div>
<script>
function log(msg, err=false) {
  const el = document.getElementById('log');
  const line = document.createElement('span');
  if (err) line.className = 'err';
  line.textContent = msg + '\\n';
  el.appendChild(line);
  el.scrollTop = el.scrollHeight;
}
function clearLog() { document.getElementById('log').innerHTML = ''; }

async function onXmlChange(input) {
  if (!input.files.length) return;
  const fd = new FormData();
  fd.append('xml', input.files[0]);
  clearLog(); log('Membaca XML…');
  try {
    const res  = await fetch('/api/preview', { method: 'POST', body: fd });
    const data = await res.json();
    if (data.error) { log('ERROR: ' + data.error, true); return; }
    document.getElementById('inv-count').textContent  = data.invoices;
    document.getElementById('item-count').textContent = data.items;
    data.preview.forEach(p => log(p));
  } catch(e) { log('Gagal: ' + e, true); }
}

async function doConvert() {
  const xmlFile  = document.getElementById('xml-file').files[0];
  const tmplFile = document.getElementById('tmpl-file').files[0];
  if (!xmlFile)  { alert('Pilih file XML Coretax.'); return; }
  if (!tmplFile) { alert('Pilih file template XLSX PajakExpress.'); return; }
  const btn = document.getElementById('conv-btn');
  btn.disabled = true; btn.textContent = 'SEDANG MENGONVERSI…';
  clearLog(); log('Memproses ' + xmlFile.name + '…');
  const fd = new FormData();
  fd.append('xml', xmlFile);
  fd.append('template', tmplFile);
  try {
    const res = await fetch('/api/convert', { method: 'POST', body: fd });
    if (!res.ok) {
      const err = await res.json();
      log('ERROR: ' + (err.error || res.statusText), true);
    } else {
      const blob = await res.blob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement('a');
      const name = xmlFile.name.replace(/\\.xml$/i,'') + '_converted.xlsx';
      a.href = url; a.download = name; a.click();
      URL.revokeObjectURL(url);
      log('✓ Selesai! File diunduh: ' + name);
      const banner = document.getElementById('success-banner');
      banner.textContent = '✓ Konversi berhasil! ' + name + ' sudah diunduh.';
      banner.style.display = 'block';
    }
  } catch(e) { log('Gagal: ' + e, true); }
  btn.disabled = false; btn.textContent = 'KONVERSI SEKARANG';
}
</script>
</body>
</html>"""


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/api/preview", methods=["POST"])
def preview():
    try:
        invoices    = parse_xml_bytes(request.files["xml"].read())
        total_items = sum(len(g) for _, g in invoices)
        preview     = [f"[{i}] {inv['nama_pembeli']} — {inv['date']} — {len(goods)} item(s)"
                       for i, (inv, goods) in enumerate(invoices, 1)]
        return jsonify(invoices=len(invoices), items=total_items, preview=preview)
    except Exception as e:
        return jsonify(error=str(e)), 400

@app.route("/api/convert", methods=["POST"])
def convert():
    try:
        invoices = parse_xml_bytes(request.files["xml"].read())
        out_buf  = write_xlsx_bytes(invoices, request.files["template"].read())
        return send_file(out_buf, as_attachment=True, download_name="converted.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as e:
        return jsonify(error=str(e)), 400


# Vercel calls `app` directly — no __main__ block needed
