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

COUNTRY_FIX = {"IND": "IDN"}

def fix_country(code): return COUNTRY_FIX.get(code, code)

def format_date(raw):
    try:
        return datetime.strptime(raw.strip(), "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return raw

def get_month(raw):
    try:
        return datetime.strptime(raw.strip(), "%Y-%m-%d").month
    except ValueError:
        return None

def get_year(raw):
    try:
        return datetime.strptime(raw.strip(), "%Y-%m-%d").strftime("%Y")
    except ValueError:
        return None

def _text(el, tag, default=""):
    child = el.find(tag)
    if child is None or child.text is None:
        return default
    return child.text.strip()

def _float(el, tag):
    """Return float value or None if tag is absent/empty."""
    child = el.find(tag)
    if child is None or not (child.text or "").strip():
        return None
    try:
        return float(child.text.strip())
    except ValueError:
        return None

def s(v):
    """Convert empty string to None so the cell is left blank."""
    if isinstance(v, str) and not v:
        return None
    return v


def parse_xml_bytes(xml_bytes):
    root = ET.fromstring(xml_bytes)
    seller_tin = _text(root, "TIN")
    invoices = []

    for inv in root.iter("TaxInvoice"):
        raw_date   = _text(inv, "TaxInvoiceDate")
        opt        = _text(inv, "TaxInvoiceOpt", "Normal")
        buyer_doc  = _text(inv, "BuyerDocument", "").lower()
        buyer_docno = _text(inv, "BuyerDocumentNumber")

        fg_pengganti = 1 if opt.lower() == "pengganti" else 0

        invoice = {
            "seller_tin":           seller_tin,
            "kd_jenis_transaksi":   s(_text(inv, "TrxCode")),
            "fg_pengganti":         fg_pengganti,
            # Try common Coretax XML tag names for invoice number
            "nomor_faktur":         s(_text(inv, "TaxInvoiceNo") or _text(inv, "SerialNo") or _text(inv, "InvoiceNo")),
            "masa_pajak":           get_month(raw_date),
            "tahun_pajak":          get_year(raw_date),
            "tanggal_faktur":       s(format_date(raw_date)),
            "npwp":                 s(_text(inv, "BuyerTin")),
            "nama":                 s(_text(inv, "BuyerName")),
            "alamat_lengkap":       s(_text(inv, "BuyerAdress")),
            "id_keterangan_tambah": s(_text(inv, "AddInfo")),
            "referensi":            s(_text(inv, "RefDesc")),
            "kode_dok_pendukung":   s(_text(inv, "CustomDoc")),
            # PASSPORT hanya diisi jika jenis ID adalah paspor
            "passport":             s(buyer_docno) if "paspor" in buyer_doc else None,
            "id_lain":              s(buyer_docno) if buyer_doc not in ("", "npwp", "nik", "paspor") else None,
            "kode_negara":          s(fix_country(_text(inv, "BuyerCountry"))),
            "id_tku_penjual":       s(_text(inv, "SellerIDTKU")),
            # Nomor faktur yang diganti (hanya ada pada pengganti)
            "nomor_faktur_diganti": s(_text(inv, "ReplacedTaxInvoiceNo") or _text(inv, "OriginalInvoiceNo")) if fg_pengganti else None,
            "email":                s(_text(inv, "BuyerEmail")),
            "keterangan1":          s(_text(inv, "FacilityStamp")),
        }

        goods = []
        for gs in inv.iter("GoodService"):
            harga = _float(gs, "Price")
            qty   = _float(gs, "Qty")
            harga_total = (harga * qty) if (harga is not None and qty is not None) else None
            goods.append({
                "kode_objek":     s(_text(gs, "Code")),
                "nama":           s(_text(gs, "Name")),
                "harga_satuan":   harga,
                "jumlah_barang":  qty,
                "harga_total":    harga_total,
                "diskon":         _float(gs, "TotalDiscount"),
                "dpp":            _float(gs, "TaxBase"),
                "ppn":            _float(gs, "VAT"),
                "tarif_ppnbm":    _float(gs, "STLGRate"),
                "ppnbm":          _float(gs, "STLG"),
                "brgjasa":        s(_text(gs, "Opt")),
                "satuanbrgjasa":  s(_text(gs, "Unit")),
                "dpp_nilai_lain": _float(gs, "OtherTaxBase"),
            })

        # Hitung total DPP / PPN / PPnBM untuk baris FK
        def safe_sum(key):
            vals = [g[key] for g in goods if g[key] is not None]
            return sum(vals) if vals else None

        invoice["jumlah_dpp"]   = safe_sum("dpp")
        invoice["jumlah_ppn"]   = safe_sum("ppn")
        invoice["jumlah_ppnbm"] = safe_sum("ppnbm")

        invoices.append((invoice, goods))

    return invoices


def write_xlsx_bytes(invoices, template_bytes):
    """
    Tulis data ke template baru.

    Struktur sheet DATA:
      Row 1 — header kolom FK  (faktur header)
      Row 2 — header kolom LT  (detail alamat penjual, tidak dipakai)
      Row 3 — header kolom OF  (objek faktur / barang-jasa)
      Row 4+ — data: satu baris FK diikuti satu atau lebih baris OF per faktur
    """
    buf = io.BytesIO(template_bytes)
    wb  = load_workbook(buf)
    ws  = wb["DATA"]

    # Hapus data lama mulai baris 4 ke bawah
    if ws.max_row and ws.max_row >= 4:
        ws.delete_rows(4, ws.max_row - 3)

    data_row = 4
    for inv, goods in invoices:
        # ── Baris FK ──────────────────────────────────────────────────────
        fk_row = [
            "FK",                          # Col 1  – penanda tipe baris
            inv["kd_jenis_transaksi"],     # Col 2  – KD_JENIS_TRANSAKSI
            inv["fg_pengganti"],           # Col 3  – FG_PENGGANTI
            inv["nomor_faktur"],           # Col 4  – NOMOR_FAKTUR
            inv["masa_pajak"],             # Col 5  – MASA_PAJAK
            inv["tahun_pajak"],            # Col 6  – TAHUN_PAJAK
            inv["tanggal_faktur"],         # Col 7  – TANGGAL_FAKTUR
            inv["npwp"],                   # Col 8  – NPWP
            inv["nama"],                   # Col 9  – NAMA
            inv["alamat_lengkap"],         # Col 10 – ALAMAT_LENGKAP
            inv["jumlah_dpp"],             # Col 11 – JUMLAH_DPP
            inv["jumlah_ppn"],             # Col 12 – JUMLAH_PPN
            inv["jumlah_ppnbm"],           # Col 13 – JUMLAH_PPNBM
            inv["id_keterangan_tambah"],   # Col 14 – ID_KETERANGAN_TAMBAH
            None,                          # Col 15 – FG_UANG_MUKA       (tidak ada di XML)
            None,                          # Col 16 – UANG_MUKA_DPP      (tidak ada di XML)
            None,                          # Col 17 – UANG_MUKA_PPN      (tidak ada di XML)
            None,                          # Col 18 – UANG_MUKA_PPNBM    (tidak ada di XML)
            None,                          # Col 19 – UANG_MUKA_DPP_LAIN (tidak ada di XML)
            inv["referensi"],              # Col 20 – REFERENSI
            inv["kode_dok_pendukung"],     # Col 21 – KODE_DOKUMEN_PENDUKUNG
            None,                          # Col 22 – NOMOR_FAKTUR_UANG_MUKA (tidak ada di XML)
            inv["passport"],               # Col 23 – PASSPORT
            inv["id_lain"],                # Col 24 – ID_LAIN
            inv["kode_negara"],            # Col 25 – KODE_NEGARA
            inv["id_tku_penjual"],         # Col 26 – ID_TKU_PENJUAL
            inv["nomor_faktur_diganti"],   # Col 27 – NOMOR_FAKTUR_DIGANTI
            inv["email"],                  # Col 28 – EMAIL
            inv["keterangan1"],            # Col 29 – KETERANGAN1
            None,                          # Col 30 – KETERANGAN2 (tidak ada di XML)
            None,                          # Col 31 – KETERANGAN3 (tidak ada di XML)
            None,                          # Col 32 – KETERANGAN4 (tidak ada di XML)
            None,                          # Col 33 – KETERANGAN5 (tidak ada di XML)
        ]
        for col, val in enumerate(fk_row, 1):
            ws.cell(row=data_row, column=col).value = val
        data_row += 1

        # ── Baris OF (satu per barang/jasa) ───────────────────────────────
        for gs in goods:
            of_row = [
                "OF",                      # Col 1  – penanda tipe baris
                gs["kode_objek"],          # Col 2  – KODE_OBJEK
                gs["nama"],                # Col 3  – NAMA
                gs["harga_satuan"],        # Col 4  – HARGA_SATUAN
                gs["jumlah_barang"],       # Col 5  – JUMLAH_BARANG
                gs["harga_total"],         # Col 6  – HARGA_TOTAL
                gs["diskon"],              # Col 7  – DISKON
                gs["dpp"],                 # Col 8  – DPP
                gs["ppn"],                 # Col 9  – PPN
                gs["tarif_ppnbm"],         # Col 10 – TARIF_PPNBM
                gs["ppnbm"],               # Col 11 – PPNBM
                gs["brgjasa"],             # Col 12 – BRGJASA
                gs["satuanbrgjasa"],       # Col 13 – SATUANBRGJASA
                gs["dpp_nilai_lain"],      # Col 14 – DPP_NILAI_LAIN
            ]
            for col, val in enumerate(of_row, 1):
                ws.cell(row=data_row, column=col).value = val
            data_row += 1

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
    <p class="sub-h">Konversi XML Coretax ke template ImporPajakKeluaran (CSV) terbaru</p>
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
      <label>Template XLSX (ImporPajakKeluaran_CSV_) *</label>
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
  <footer>Coretax DLP XML → ImporPajakKeluaran XLSX &nbsp;|</footer>
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
  if (!tmplFile) { alert('Pilih file template XLSX ImporPajakKeluaran.'); return; }
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
        preview     = [
            f"[{i}] {inv['nama'] or '—'} — {inv['tanggal_faktur'] or '—'} — {len(goods)} item(s)"
            for i, (inv, goods) in enumerate(invoices, 1)
        ]
        return jsonify(invoices=len(invoices), items=total_items, preview=preview)
    except Exception as e:
        return jsonify(error=str(e)), 400

@app.route("/api/convert", methods=["POST"])
def convert():
    try:
        invoices = parse_xml_bytes(request.files["xml"].read())
        out_buf  = write_xlsx_bytes(invoices, request.files["template"].read())
        return send_file(
            out_buf,
            as_attachment=True,
            download_name="converted.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        return jsonify(error=str(e)), 400


# Vercel calls `app` directly — no __main__ block needed
