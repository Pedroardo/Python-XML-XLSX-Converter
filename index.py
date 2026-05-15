"""
api/index.py — Vercel Serverless Function
Coretax DLP XML  →  ImporPajakKeluaran XLSX / CSV (PajakExpress)
"""

from flask import Flask, request, send_file, jsonify, render_template_string
import io, csv
from datetime import datetime
from xml.etree import ElementTree as ET
from openpyxl import load_workbook

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

# IND adalah kode lama / salah; IDN adalah ISO 3166-1 alpha-3 yang benar
COUNTRY_FIX = {"IND": "IDN"}

# Coretax <Opt> → BRGJASA yang dipakai template PajakExpress
OPT_MAP = {
    "A": "GOODS",     # Barang
    "B": "GOODS",
    "J": "SERVICES",  # Jasa
    "S": "SERVICES",
    "GOODS":    "GOODS",
    "SERVICES": "SERVICES",
}

# BuyerDocument yang menandakan pembeli punya NPWP/NIK lokal → KODE_NEGARA kosong
DOMESTIC_DOC = {"tin", "npwp", "nik", ""}


def _text(el, tag):
    """Teks child tag; None jika tag tidak ada atau kosong (termasuk <Tag />)."""
    child = el.find(tag)
    if child is None:
        return None
    t = (child.text or "").strip()
    return t if t else None


def _num(el, tag):
    """Float dari child tag; None jika tidak ada / kosong."""
    raw = _text(el, tag)
    if raw is None:
        return None
    try:
        return float(raw)
    except ValueError:
        return None


def to_int_or_float(v):
    """Kembalikan int jika bilangan bulat, float jika desimal, None jika None."""
    if v is None:
        return None
    return int(v) if v == int(v) else v


def fmt_str(v):
    """Format angka jadi string (tanpa desimal jika bulat). None → None."""
    if v is None:
        return None
    return str(int(v)) if v == int(v) else str(v)


def format_date(raw):
    try:
        return datetime.strptime(raw.strip(), "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return raw or None


def get_month(raw):       # integer, bukan string '04'
    try:
        return datetime.strptime(raw.strip(), "%Y-%m-%d").month
    except Exception:
        return None


def get_year(raw):        # string '2026'
    try:
        return datetime.strptime(raw.strip(), "%Y-%m-%d").strftime("%Y")
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────────────────────
# XML Parsing
# ─────────────────────────────────────────────────────────────────────────────

def parse_xml_bytes(xml_bytes):
    root = ET.fromstring(xml_bytes)
    seller_tin = _text(root, "TIN")
    invoices   = []

    for inv in root.iter("TaxInvoice"):
        raw_date   = _text(inv, "TaxInvoiceDate") or ""
        opt        = (_text(inv, "TaxInvoiceOpt") or "Normal").strip()
        buyer_doc  = (_text(inv, "BuyerDocument") or "").strip().lower()
        buyer_docno = _text(inv, "BuyerDocumentNumber")

        fg_pengganti = 1 if opt.lower() == "pengganti" else 0

        # KODE_NEGARA hanya diisi jika pembeli tidak punya NPWP/NIK domestik
        is_domestic  = buyer_doc in DOMESTIC_DOC
        raw_country  = _text(inv, "BuyerCountry") or ""
        kode_negara  = None if is_domestic else COUNTRY_FIX.get(raw_country.upper(), raw_country) or None

        invoice = {
            "seller_tin":           seller_tin,
            # FK columns
            "kd_jenis_transaksi":   _text(inv, "TrxCode"),           # str '04'
            "fg_pengganti":         fg_pengganti,                     # int 0/1
            # nomor_faktur diisi oleh generator, bukan dari XML
            "masa_pajak":           get_month(raw_date),              # int 4
            "tahun_pajak":          get_year(raw_date),               # str '2026'
            "tanggal_faktur":       format_date(raw_date),            # str 'dd/MM/yyyy'
            "npwp":                 _text(inv, "BuyerTin"),           # str (leading zero aman)
            "nama":                 _text(inv, "BuyerName"),
            "alamat_lengkap":       _text(inv, "BuyerAdress"),
            "id_keterangan_tambah": _text(inv, "AddInfo"),            # None jika <AddInfo />
            "referensi":            _text(inv, "RefDesc"),
            "kode_dok_pendukung":   _text(inv, "CustomDoc"),          # None jika <CustomDoc />
            "passport":             buyer_docno if "paspor" in buyer_doc else None,
            "id_lain":              buyer_docno if buyer_doc not in DOMESTIC_DOC and "paspor" not in buyer_doc else None,
            "kode_negara":          kode_negara,
            "id_tku_penjual":       _text(inv, "SellerIDTKU"),
            "nomor_faktur_diganti": _text(inv, "ReplacedTaxInvoiceNo") if fg_pengganti else None,
            "email":                _text(inv, "BuyerEmail"),         # None jika <BuyerEmail />
            "keterangan1":          _text(inv, "FacilityStamp"),      # None jika <FacilityStamp />
        }

        goods = []
        for gs in inv.iter("GoodService"):
            price = _num(gs, "Price")
            qty   = _num(gs, "Qty")
            harga_total = (price * qty) if (price is not None and qty is not None) else None

            goods.append({
                "kode_objek":     _text(gs, "Code"),                  # str, bukan int
                "nama":           _text(gs, "Name"),
                "harga_satuan":   to_int_or_float(price),             # int jika bulat
                "jumlah_barang":  to_int_or_float(qty),
                "harga_total":    to_int_or_float(harga_total),
                "diskon":         to_int_or_float(_num(gs, "TotalDiscount")),
                "dpp":            to_int_or_float(_num(gs, "TaxBase")),
                "ppn":            to_int_or_float(_num(gs, "VAT")),
                "tarif_ppnbm":    fmt_str(_num(gs, "STLGRate")),      # str '0' sesuai template
                "ppnbm":          fmt_str(_num(gs, "STLG")),          # str '0' sesuai template
                "brgjasa":        OPT_MAP.get((_text(gs, "Opt") or "").strip().upper(), "GOODS"),
                "satuanbrgjasa":  _text(gs, "Unit"),
                "dpp_nilai_lain": to_int_or_float(_num(gs, "OtherTaxBase")),
            })

        # Total untuk baris FK — disimpan sebagai string sesuai format template
        def total_str(key):
            vals = [g[key] for g in goods if g[key] is not None]
            if not vals:
                return None
            s = sum(float(v) for v in vals)
            return str(int(s)) if s == int(s) else str(s)

        invoice["jumlah_dpp"]   = total_str("dpp")    # str
        invoice["jumlah_ppn"]   = total_str("ppn")    # str
        invoice["jumlah_ppnbm"] = total_str("ppnbm")  # str

        invoices.append((invoice, goods))

    return invoices


# ─────────────────────────────────────────────────────────────────────────────
# NOMOR_FAKTUR Generator
# ─────────────────────────────────────────────────────────────────────────────

def gen_nomor(inv, seq, prefix, start):
    """
    Buat NOMOR_FAKTUR unik.
      • Jika prefix diisi  → {prefix}{nomor 5 digit}   cth: INV00001
      • Jika prefix kosong → {YYYY}{MM}{nomor 5 digit}  cth: 20260400001
    seq   : 0-based index faktur dalam batch
    start : nomor urut awal yang dipilih user
    """
    num = str(start + seq).zfill(5)
    if prefix:
        return f"{prefix}{num}"
    tahun = inv.get("tahun_pajak") or ""
    masa  = str(inv.get("masa_pajak") or "").zfill(2)
    return f"{tahun}{masa}{num}"


# ─────────────────────────────────────────────────────────────────────────────
# Row Builder  (shared antara XLSX dan CSV)
# ─────────────────────────────────────────────────────────────────────────────

def build_data_rows(invoices, prefix, start):
    """
    Kembalikan list baris data.
    FK row  = 33 kolom (sesuai header Row 1 template)
    OF row  = 14 kolom diisi + 19 None (total 33, supaya CSV seragam)
    """
    rows = []
    for seq, (inv, goods) in enumerate(invoices):
        nomor = gen_nomor(inv, seq, prefix, start)

        fk = [
            "FK",                           # C1  – tipe baris
            inv["kd_jenis_transaksi"],       # C2  – KD_JENIS_TRANSAKSI
            inv["fg_pengganti"],             # C3  – FG_PENGGANTI          int
            nomor,                           # C4  – NOMOR_FAKTUR           str
            inv["masa_pajak"],               # C5  – MASA_PAJAK             int
            inv["tahun_pajak"],              # C6  – TAHUN_PAJAK            str
            inv["tanggal_faktur"],           # C7  – TANGGAL_FAKTUR         str dd/MM/yyyy
            inv["npwp"],                     # C8  – NPWP                   str
            inv["nama"],                     # C9  – NAMA
            inv["alamat_lengkap"],           # C10 – ALAMAT_LENGKAP
            inv["jumlah_dpp"],               # C11 – JUMLAH_DPP             str
            inv["jumlah_ppn"],               # C12 – JUMLAH_PPN             str
            inv["jumlah_ppnbm"],             # C13 – JUMLAH_PPNBM           str
            inv["id_keterangan_tambah"],     # C14 – ID_KETERANGAN_TAMBAH
            None,                            # C15 – FG_UANG_MUKA           tidak ada di XML
            None,                            # C16 – UANG_MUKA_DPP          tidak ada di XML
            None,                            # C17 – UANG_MUKA_PPN          tidak ada di XML
            None,                            # C18 – UANG_MUKA_PPNBM        tidak ada di XML
            None,                            # C19 – UANG_MUKA_DPP_LAIN     tidak ada di XML
            inv["referensi"],                # C20 – REFERENSI
            inv["kode_dok_pendukung"],       # C21 – KODE_DOKUMEN_PENDUKUNG
            None,                            # C22 – NOMOR_FAKTUR_UANG_MUKA  tidak ada di XML
            inv["passport"],                 # C23 – PASSPORT
            inv["id_lain"],                  # C24 – ID_LAIN
            inv["kode_negara"],              # C25 – KODE_NEGARA
            inv["id_tku_penjual"],           # C26 – ID_TKU_PENJUAL
            inv["nomor_faktur_diganti"],     # C27 – NOMOR_FAKTUR_DIGANTI
            inv["email"],                    # C28 – EMAIL
            inv["keterangan1"],              # C29 – KETERANGAN1
            None,                            # C30 – KETERANGAN2             tidak ada di XML
            None,                            # C31 – KETERANGAN3             tidak ada di XML
            None,                            # C32 – KETERANGAN4             tidak ada di XML
            None,                            # C33 – KETERANGAN5             tidak ada di XML
        ]
        rows.append(fk)

        for gs in goods:
            of = [
                "OF",                        # C1  – tipe baris
                gs["kode_objek"],            # C2  – KODE_OBJEK             str
                gs["nama"],                  # C3  – NAMA
                gs["harga_satuan"],          # C4  – HARGA_SATUAN           int/float
                gs["jumlah_barang"],         # C5  – JUMLAH_BARANG          int/float
                gs["harga_total"],           # C6  – HARGA_TOTAL            int/float
                gs["diskon"],                # C7  – DISKON                 int/float
                gs["dpp"],                   # C8  – DPP                    int/float
                gs["ppn"],                   # C9  – PPN                    int/float
                gs["tarif_ppnbm"],           # C10 – TARIF_PPNBM            str
                gs["ppnbm"],                 # C11 – PPNBM                  str
                gs["brgjasa"],               # C12 – BRGJASA                str GOODS/SERVICES
                gs["satuanbrgjasa"],         # C13 – SATUANBRGJASA          str UM.xxxx
                gs["dpp_nilai_lain"],        # C14 – DPP_NILAI_LAIN         int/float
            ] + [None] * 19                  # C15-C33 kosong (padded)
            rows.append(of)

    return rows


# ─────────────────────────────────────────────────────────────────────────────
# XLSX Writer
# ─────────────────────────────────────────────────────────────────────────────

def write_xlsx(invoices, template_bytes, prefix, start):
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb["DATA"]

    # Hapus data lama (baris 4+), pertahankan 3 baris header template
    if ws.max_row and ws.max_row >= 4:
        ws.delete_rows(4, ws.max_row - 3)

    for row in build_data_rows(invoices, prefix, start):
        ws.append(row)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# ─────────────────────────────────────────────────────────────────────────────
# CSV Writer
# ─────────────────────────────────────────────────────────────────────────────

# Baris header CSV  (mirror rows 1-3 dari template DATA sheet)
_H_FK = [
    "FK","KD_JENIS_TRANSAKSI","FG_PENGGANTI","NOMOR_FAKTUR","MASA_PAJAK",
    "TAHUN_PAJAK","TANGGAL_FAKTUR","NPWP","NAMA","ALAMAT_LENGKAP",
    "JUMLAH_DPP","JUMLAH_PPN","JUMLAH_PPNBM","ID_KETERANGAN_TAMBAH",
    "FG_UANG_MUKA","UANG_MUKA_DPP","UANG_MUKA_PPN","UANG_MUKA_PPNBM",
    "UANG_MUKA_DPP_LAIN","REFERENSI","KODE_DOKUMEN_PENDUKUNG",
    "NOMOR_FAKTUR_UANG_MUKA","PASSPORT","ID_LAIN","KODE_NEGARA",
    "ID_TKU_PENJUAL","NOMOR_FAKTUR_DIGANTI","EMAIL",
    "KETERANGAN1","KETERANGAN2","KETERANGAN3","KETERANGAN4","KETERANGAN5",
]
_H_LT = (["LT","NPWP","NAMA","JALAN","BLOK","NOMOR","RT","RW",
           "KECAMATAN","KELURAHAN","KABUPATEN","PROPINSI","KODE_POS","NOMOR_TELEPON"]
          + [""] * 19)
_H_OF = (["OF","KODE_OBJEK","NAMA","HARGA_SATUAN","JUMLAH_BARANG","HARGA_TOTAL",
           "DISKON","DPP","PPN","TARIF_PPNBM","PPNBM","BRGJASA","SATUANBRGJASA",
           "DPP_NILAI_LAIN"]
         + [""] * 19)


def write_csv(invoices, prefix, start):
    buf = io.StringIO()
    w   = csv.writer(buf, lineterminator="\r\n")

    w.writerow(_H_FK)
    w.writerow(_H_LT)
    w.writerow(_H_OF)

    for row in build_data_rows(invoices, prefix, start):
        w.writerow(["" if v is None else v for v in row])

    # UTF-8 BOM agar Excel tidak salah baca karakter khusus
    return io.BytesIO(buf.getvalue().encode("utf-8-sig"))


# ─────────────────────────────────────────────────────────────────────────────
# HTML
# ─────────────────────────────────────────────────────────────────────────────

HTML = """<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Coretax → PajakExpress</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');
:root{
  --bg:#080b10;--card:#0e1318;--border:#1e2830;
  --accent:#00d4ff;--green:#00ff9d;--yellow:#ffd166;
  --text:#c8d8e8;--sub:#4a6070;--danger:#ff4d6d;
}
*{box-sizing:border-box;margin:0;padding:0;}
body{background:var(--bg);color:var(--text);font-family:'IBM Plex Sans',sans-serif;
     min-height:100vh;display:flex;flex-direction:column;align-items:center;
     padding:40px 16px 60px;}
body::before{content:'';position:fixed;inset:0;pointer-events:none;z-index:0;
  background-image:linear-gradient(rgba(0,212,255,.04) 1px,transparent 1px),
                   linear-gradient(90deg,rgba(0,212,255,.04) 1px,transparent 1px);
  background-size:40px 40px;}
.wrap{position:relative;z-index:1;width:100%;max-width:600px;}

/* Header */
header{text-align:center;margin-bottom:32px;}
.logo{font-size:11px;color:var(--accent);letter-spacing:4px;
      font-family:'IBM Plex Mono',monospace;margin-bottom:10px;}
h1{font-size:24px;font-weight:600;color:#fff;}
h1 span{color:var(--accent);}
.sub-h{font-size:12px;color:var(--sub);margin-top:5px;}

/* Cards */
.card{background:var(--card);border:1px solid var(--border);
      border-radius:6px;padding:20px;margin-bottom:12px;}
.section-label{font-family:'IBM Plex Mono',monospace;font-size:10px;
               color:var(--sub);letter-spacing:2px;margin-bottom:14px;}
.field+.field{margin-top:12px;}
label{display:block;font-size:11px;color:var(--sub);margin-bottom:5px;}
.file-input,.text-input{
  width:100%;background:#050709;border:1px solid var(--border);border-radius:4px;
  color:var(--text);font-family:'IBM Plex Mono',monospace;font-size:11px;
  padding:9px 12px;outline:none;cursor:pointer;transition:border-color .2s;}
.file-input:hover,.text-input:hover,.text-input:focus{border-color:var(--accent);}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:10px;}
.hint{font-size:11px;color:var(--sub);margin-top:6px;line-height:1.5;}
.hint code{color:var(--accent);}

/* Format toggle */
.fmt-toggle{display:flex;gap:8px;}
.fmt-btn{flex:1;background:#050709;border:1px solid var(--border);border-radius:4px;
          color:var(--sub);font-family:'IBM Plex Mono',monospace;font-size:12px;
          padding:10px;cursor:pointer;transition:all .2s;text-align:center;letter-spacing:1px;}
.fmt-btn.active{border-color:var(--accent);color:var(--accent);background:rgba(0,212,255,.06);}

/* Stats */
.stats{display:flex;gap:10px;margin-bottom:12px;}
.stat{flex:1;background:#050709;border:1px solid var(--border);
      border-radius:4px;padding:12px;text-align:center;}
.stat-num{font-size:26px;font-weight:600;color:var(--accent);
          font-family:'IBM Plex Mono',monospace;}
.stat-num.g{color:var(--green);}

/* Log */
.log{background:#020304;border:1px solid var(--border);border-radius:4px;
     font-family:'IBM Plex Mono',monospace;font-size:11px;color:var(--green);
     padding:12px;height:130px;overflow-y:auto;white-space:pre-wrap;word-break:break-all;}
.log .err{color:var(--danger);}
.log .warn{color:var(--yellow);}

/* Button */
.btn{width:100%;background:var(--accent);color:#000;border:none;border-radius:4px;
     font-family:'IBM Plex Mono',monospace;font-size:13px;font-weight:600;
     letter-spacing:2px;padding:15px;cursor:pointer;transition:opacity .2s,transform .1s;}
.btn:hover{opacity:.88;}
.btn:active{transform:scale(.99);}
.btn:disabled{opacity:.35;cursor:not-allowed;}

/* Banner */
.banner{display:none;background:rgba(0,255,157,.08);border:1px solid var(--green);
        border-radius:4px;padding:13px;color:var(--green);font-size:13px;
        margin-bottom:12px;text-align:center;}

footer{margin-top:36px;font-size:11px;color:var(--sub);text-align:center;}
</style>
</head>
<body>
<div class="wrap">

<header>
  <div class="logo">PAJAK CONVERTER</div>
  <h1>XML <span>→</span> PajakExpress</h1>
  <p class="sub-h">Coretax DLP XML → ImporPajakKeluaran XLSX / CSV</p>
</header>

<div id="banner" class="banner"></div>

<!-- 01 FILES -->
<div class="card">
  <div class="section-label">01 / INPUT FILES</div>
  <div class="field">
    <label>File XML Coretax *</label>
    <input class="file-input" type="file" id="xml-file" accept=".xml" onchange="onXmlChange(this)">
  </div>
  <div class="field" id="tmpl-wrap">
    <label>Template XLSX (ImporPajakKeluaran_CSV_) — wajib untuk output XLSX</label>
    <input class="file-input" type="file" id="tmpl-file" accept=".xlsx">
  </div>
</div>

<!-- 02 NOMOR FAKTUR -->
<div class="card">
  <div class="section-label">02 / NOMOR FAKTUR <span style="color:var(--danger)">*</span></div>
  <p class="hint" style="margin-bottom:12px;">
    NOMOR_FAKTUR tidak ada di XML Coretax — akan di-generate otomatis.<br>
    Default: <code>YYYYMM</code> + 5 digit urutan &nbsp;→&nbsp; <code id="nomor-ex">20260400001</code>
  </p>
  <div class="row2">
    <div class="field">
      <label>Prefix kustom (opsional)</label>
      <input class="text-input" type="text" id="prefix" placeholder="cth: INV atau kosongkan" oninput="updatePreview()">
    </div>
    <div class="field">
      <label>Nomor urut awal</label>
      <input class="text-input" type="number" id="start-num" value="1" min="1" oninput="updatePreview()">
    </div>
  </div>
</div>

<!-- 03 FORMAT -->
<div class="card">
  <div class="section-label">03 / FORMAT OUTPUT</div>
  <div class="fmt-toggle">
    <div class="fmt-btn active" id="btn-xlsx" onclick="setFmt('xlsx')">📊 &nbsp;XLSX</div>
    <div class="fmt-btn"        id="btn-csv"  onclick="setFmt('csv')"> 📄 &nbsp;CSV</div>
  </div>
  <p class="hint" style="margin-top:10px;" id="fmt-hint">
    XLSX: isi template ImporPajakKeluaran (butuh upload template di atas).
  </p>
</div>

<!-- 04 PREVIEW -->
<div class="card">
  <div class="section-label">04 / PREVIEW</div>
  <div class="stats">
    <div class="stat"><div class="stat-num" id="inv-count">—</div><div style="font-size:10px;color:var(--sub);margin-top:2px">Faktur (FK)</div></div>
    <div class="stat"><div class="stat-num g" id="item-count">—</div><div style="font-size:10px;color:var(--sub);margin-top:2px">Item (OF)</div></div>
  </div>
  <div class="log" id="log">Pilih file XML untuk melihat preview…</div>
</div>

<button class="btn" id="conv-btn" onclick="doConvert()">▶ &nbsp;CONVERT &amp; DOWNLOAD</button>
<footer>Coretax DLP XML → ImporPajakKeluaran &nbsp;|&nbsp; XLSX &amp; CSV</footer>
</div>

<script>
let fmt = 'xlsx';

function setFmt(f) {
  fmt = f;
  document.getElementById('btn-xlsx').classList.toggle('active', f==='xlsx');
  document.getElementById('btn-csv').classList.toggle('active',  f==='csv');
  document.getElementById('tmpl-wrap').style.display = f==='xlsx' ? '' : 'none';
  document.getElementById('fmt-hint').textContent = f === 'xlsx'
    ? 'XLSX: isi template ImporPajakKeluaran (butuh upload template di atas).'
    : 'CSV: UTF-8 BOM, flat file, langsung bisa diupload ke PajakExpress.';
}

function updatePreview() {
  const prefix = document.getElementById('prefix').value.trim();
  const start  = parseInt(document.getElementById('start-num').value) || 1;
  const num    = String(start).padStart(5,'0');
  const sample = prefix ? (prefix + num) : ('YYYYMM' + num);
  document.getElementById('nomor-ex').textContent = sample;
}
updatePreview();

function addLog(msg, cls='') {
  const el   = document.getElementById('log');
  const span = document.createElement('span');
  if (cls) span.className = cls;
  span.textContent = msg + '\n';
  el.appendChild(span);
  el.scrollTop = el.scrollHeight;
}
function clearLog() { document.getElementById('log').innerHTML = ''; }

async function onXmlChange(input) {
  if (!input.files.length) return;
  const fd = new FormData();
  fd.append('xml', input.files[0]);
  clearLog(); addLog('Membaca XML…');
  try {
    const res  = await fetch('/api/preview', {method:'POST', body:fd});
    const data = await res.json();
    if (data.error) { addLog('ERROR: ' + data.error, 'err'); return; }
    document.getElementById('inv-count').textContent  = data.invoices;
    document.getElementById('item-count').textContent = data.items;
    (data.warnings || []).forEach(w => addLog('⚠ ' + w, 'warn'));
    data.preview.forEach(p => addLog(p));
  } catch(e) { addLog('Gagal: ' + e, 'err'); }
}

async function doConvert() {
  const xmlFile  = document.getElementById('xml-file').files[0];
  const tmplFile = document.getElementById('tmpl-file').files[0];
  const prefix   = document.getElementById('prefix').value.trim();
  const startNum = parseInt(document.getElementById('start-num').value) || 1;

  if (!xmlFile) { alert('Pilih file XML Coretax terlebih dahulu.'); return; }
  if (fmt === 'xlsx' && !tmplFile) {
    alert('Untuk output XLSX, upload dulu file template ImporPajakKeluaran_CSV_.xlsx.\nAtau ganti format ke CSV.'); return;
  }

  const btn = document.getElementById('conv-btn');
  btn.disabled = true; btn.textContent = 'SEDANG MENGONVERSI…';
  clearLog(); addLog('Memproses ' + xmlFile.name + '…');

  const fd = new FormData();
  fd.append('xml',       xmlFile);
  fd.append('format',    fmt);
  fd.append('prefix',    prefix);
  fd.append('start_num', startNum);
  if (tmplFile) fd.append('template', tmplFile);

  try {
    const res = await fetch('/api/convert', {method:'POST', body:fd});
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      addLog('ERROR: ' + (err.error || res.statusText), 'err');
    } else {
      const blob = await res.blob();
      const ext  = fmt === 'csv' ? '.csv' : '.xlsx';
      const name = xmlFile.name.replace(/\\.xml$/i,'') + '_converted' + ext;
      const a    = document.createElement('a');
      a.href = URL.createObjectURL(blob); a.download = name; a.click();
      URL.revokeObjectURL(a.href);
      addLog('✓ Selesai — ' + name + ' sudah diunduh.');
      const b = document.getElementById('banner');
      b.textContent = '✓ Konversi berhasil! File: ' + name;
      b.style.display = 'block';
    }
  } catch(e) { addLog('Gagal: ' + e, 'err'); }
  btn.disabled = false; btn.textContent = '▶  CONVERT & DOWNLOAD';
}
</script>
</body>
</html>"""


# ─────────────────────────────────────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template_string(HTML)


@app.route("/api/preview", methods=["POST"])
def preview():
    try:
        invoices    = parse_xml_bytes(request.files["xml"].read())
        total_items = sum(len(g) for _, g in invoices)
        warnings    = []
        lines       = []

        for i, (inv, goods) in enumerate(invoices, 1):
            if not inv.get("npwp"):
                warnings.append(f"Faktur #{i}: NPWP pembeli kosong")
            if not inv.get("kd_jenis_transaksi"):
                warnings.append(f"Faktur #{i}: TrxCode tidak ditemukan di XML")
            lines.append(
                f"[{i}] {inv['nama'] or '—'}  |  "
                f"{inv['tanggal_faktur'] or '—'}  |  "
                f"TrxCode={inv['kd_jenis_transaksi'] or '—'}  |  "
                f"{len(goods)} item"
            )

        return jsonify(invoices=len(invoices), items=total_items,
                       preview=lines, warnings=warnings)
    except Exception as e:
        return jsonify(error=str(e)), 400


@app.route("/api/convert", methods=["POST"])
def convert():
    try:
        invoices = parse_xml_bytes(request.files["xml"].read())
        fmt      = request.form.get("format", "xlsx").lower()
        prefix   = request.form.get("prefix", "").strip()
        start    = int(request.form.get("start_num", 1) or 1)

        if fmt == "csv":
            buf      = write_csv(invoices, prefix, start)
            mime     = "text/csv"
            dl_name  = "converted.csv"
        else:
            tmpl = request.files.get("template")
            if not tmpl:
                return jsonify(error="Template XLSX wajib disertakan untuk output XLSX."), 400
            buf     = write_xlsx(invoices, tmpl.read(), prefix, start)
            mime    = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            dl_name = "converted.xlsx"

        return send_file(buf, as_attachment=True, download_name=dl_name, mimetype=mime)

    except Exception as e:
        return jsonify(error=str(e)), 400


# Vercel memanggil `app` langsung — tidak perlu blok __main__
