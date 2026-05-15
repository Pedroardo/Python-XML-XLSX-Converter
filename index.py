"""
api/index.py  —  Coretax DLP XML  →  ImporPajakKeluaran XLSX / CSV
Vercel Serverless Function (Flask).
Untuk local dev: python index.py
"""

from flask import Flask, request, send_file, jsonify, Response
import io, csv, math
from datetime import datetime
from xml.etree import ElementTree as ET
from openpyxl import load_workbook

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024 * 1024


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

COUNTRY_FIX = {"IND": "IDN"}

OPT_MAP = {
    "A": "GOODS", "B": "GOODS",
    "J": "SERVICES", "S": "SERVICES",
    "GOODS": "GOODS", "SERVICES": "SERVICES",
}

DOMESTIC_DOC = {"tin", "npwp", "nik", ""}


def _text(el, tag):
    child = el.find(tag)
    if child is None:
        return None
    t = (child.text or "").strip()
    return t if t else None


def _num(el, tag):
    raw = _text(el, tag)
    if raw is None:
        return None
    try:
        return float(raw)
    except ValueError:
        return None


def _round_no_decimal(v):
    """Round to integer: fraction >= .5 rounds up, fraction < .5 rounds down.
    e.g. 5.5 -> 6, 5.48 -> 5, 9.55 -> 10, 493487.5 -> 493488"""
    if v is None:
        return None
    frac = v - math.floor(v)
    return math.ceil(v) if frac >= 0.5 else math.floor(v)


def to_int_or_float(v):
    if v is None:
        return None
    return _round_no_decimal(v)


def fmt_str(v):
    if v is None:
        return None
    return str(_round_no_decimal(v))


def format_date(raw):
    try:
        return datetime.strptime(raw.strip(), "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return raw or None


def get_month(raw):
    try:
        return datetime.strptime(raw.strip(), "%Y-%m-%d").month
    except Exception:
        return None


def get_year(raw):
    try:
        return datetime.strptime(raw.strip(), "%Y-%m-%d").strftime("%Y")
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────────────────────
# XML Parsing
# ─────────────────────────────────────────────────────────────────────────────

def parse_xml_bytes(xml_bytes):
    root       = ET.fromstring(xml_bytes)
    seller_tin = _text(root, "TIN")
    invoices   = []

    for inv in root.iter("TaxInvoice"):
        raw_date    = _text(inv, "TaxInvoiceDate") or ""
        opt         = (_text(inv, "TaxInvoiceOpt") or "Normal").strip()
        buyer_doc   = (_text(inv, "BuyerDocument") or "").strip().lower()
        buyer_docno = _text(inv, "BuyerDocumentNumber")

        fg_pengganti = 1 if opt.lower() == "pengganti" else 0

        is_domestic = buyer_doc in DOMESTIC_DOC
        raw_country = _text(inv, "BuyerCountry") or ""
        kode_negara = None if is_domestic else (
            COUNTRY_FIX.get(raw_country.upper(), raw_country) or None
        )

        invoice = {
            "seller_tin":           seller_tin,
            "kd_jenis_transaksi":   _text(inv, "TrxCode"),
            "fg_pengganti":         fg_pengganti,
            "masa_pajak":           get_month(raw_date),
            "tahun_pajak":          get_year(raw_date),
            "tanggal_faktur":       format_date(raw_date),
            "npwp":                 _text(inv, "BuyerTin"),
            "nama":                 _text(inv, "BuyerName"),
            "alamat_lengkap":       _text(inv, "BuyerAdress"),
            "id_keterangan_tambah": _text(inv, "AddInfo"),
            "referensi":            _text(inv, "RefDesc"),
            "kode_dok_pendukung":   _text(inv, "CustomDoc"),
            "passport":             buyer_docno if "paspor" in buyer_doc else None,
            "id_lain":              buyer_docno if (
                buyer_doc not in DOMESTIC_DOC and "paspor" not in buyer_doc
            ) else None,
            "kode_negara":          kode_negara,
            "id_tku_penjual":       _text(inv, "SellerIDTKU"),
            "nomor_faktur_diganti": _text(inv, "ReplacedTaxInvoiceNo") if fg_pengganti else None,
            "email":                _text(inv, "BuyerEmail"),
            "keterangan1":          _text(inv, "FacilityStamp"),
        }

        goods = []
        for gs in inv.iter("GoodService"):
            price    = _num(gs, "Price")
            qty      = _num(gs, "Qty")
            tax_base = _num(gs, "TaxBase")
            discount = _num(gs, "TotalDiscount")
            harga_total = (price * qty) if (price is not None and qty is not None) else None
            # DPP = TaxBase - TotalDiscount
            dpp_val = ((tax_base or 0) - (discount or 0)) if tax_base is not None else None
            goods.append({
                "kode_objek":     _text(gs, "Code"),
                "nama":           _text(gs, "Name"),
                "harga_satuan":   to_int_or_float(price),
                "jumlah_barang":  to_int_or_float(qty),
                "harga_total":    to_int_or_float(harga_total),
                "diskon":         to_int_or_float(discount),
                "dpp":            to_int_or_float(dpp_val),
                "ppn":            to_int_or_float(_num(gs, "VAT")),
                "tarif_ppnbm":    fmt_str(_num(gs, "STLGRate")),
                "ppnbm":          fmt_str(_num(gs, "STLG")),
                "brgjasa":        OPT_MAP.get((_text(gs, "Opt") or "").strip().upper(), "GOODS"),
                "satuanbrgjasa":  _text(gs, "Unit"),
                "dpp_nilai_lain": to_int_or_float(_num(gs, "OtherTaxBase")),
            })

        def total_str(key):
            vals = [g[key] for g in goods if g[key] is not None]
            if not vals:
                return None
            s = sum(float(v) for v in vals)
            return str(_round_no_decimal(s))

        invoice["jumlah_dpp"]   = total_str("dpp")
        invoice["jumlah_ppn"]   = total_str("ppn")
        invoice["jumlah_ppnbm"] = total_str("ppnbm")
        invoices.append((invoice, goods))

    return invoices


# ─────────────────────────────────────────────────────────────────────────────
# NOMOR_FAKTUR Generator
# ─────────────────────────────────────────────────────────────────────────────

def gen_nomor(inv, seq, prefix, start):
    num = str(start + seq).zfill(5)
    if prefix:
        return f"{prefix}{num}"
    tahun = inv.get("tahun_pajak") or ""
    masa  = str(inv.get("masa_pajak") or "").zfill(2)
    return f"{tahun}{masa}{num}"


# ─────────────────────────────────────────────────────────────────────────────
# Row Builder
# ─────────────────────────────────────────────────────────────────────────────

def build_data_rows(invoices, prefix, start):
    rows = []
    for seq, (inv, goods) in enumerate(invoices):
        nomor = gen_nomor(inv, seq, prefix, start)
        fk = [
            "FK",
            inv["kd_jenis_transaksi"],
            inv["fg_pengganti"],
            nomor,
            inv["masa_pajak"],
            inv["tahun_pajak"],
            inv["tanggal_faktur"],
            inv["npwp"],
            inv["nama"],
            inv["alamat_lengkap"],
            inv["jumlah_dpp"],
            inv["jumlah_ppn"],
            inv["jumlah_ppnbm"],
            inv["id_keterangan_tambah"],
            None, None, None, None, None,   # uang muka (tidak ada di XML)
            inv["referensi"],
            inv["kode_dok_pendukung"],
            None,                            # nomor_faktur_uang_muka
            inv["passport"],
            inv["id_lain"],
            inv["kode_negara"],
            inv["id_tku_penjual"],
            inv["nomor_faktur_diganti"],
            inv["email"],
            inv["keterangan1"],
            None, None, None, None,          # keterangan 2-5
        ]
        rows.append(fk)

        for gs in goods:
            of = [
                "OF",
                gs["kode_objek"],
                gs["nama"],
                gs["harga_satuan"],
                gs["jumlah_barang"],
                gs["harga_total"],
                gs["diskon"],
                gs["dpp"],
                gs["ppn"],
                gs["tarif_ppnbm"],
                gs["ppnbm"],
                gs["brgjasa"],
                gs["satuanbrgjasa"],
                gs["dpp_nilai_lain"],
            ] + [None] * 19
            rows.append(of)

    return rows


# ─────────────────────────────────────────────────────────────────────────────
# Writers
# ─────────────────────────────────────────────────────────────────────────────

def write_xlsx(invoices, template_bytes, prefix, start):
    wb = load_workbook(io.BytesIO(template_bytes))
    ws = wb["DATA"]
    if ws.max_row and ws.max_row >= 4:
        ws.delete_rows(4, ws.max_row - 3)
    for row in build_data_rows(invoices, prefix, start):
        ws.append(row)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


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
           "DPP_NILAI_LAIN"] + [""] * 19)


def write_csv(invoices, prefix, start):
    buf = io.StringIO()
    w   = csv.writer(buf, lineterminator="\r\n")
    w.writerow(_H_FK)
    w.writerow(_H_LT)
    w.writerow(_H_OF)
    for row in build_data_rows(invoices, prefix, start):
        w.writerow(["" if v is None else v for v in row])
    return io.BytesIO(buf.getvalue().encode("utf-8-sig"))


# ─────────────────────────────────────────────────────────────────────────────
# HTML  — NOTE: pakai Response() bukan render_template_string()
#         supaya Jinja2 tidak memproses CSS/JS braces
# ─────────────────────────────────────────────────────────────────────────────

HTML = r"""<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Coretax &rarr; PajakExpress</title>
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600&display=swap');
:root {
  --bg:#080b10; --card:#0e1318; --border:#1e2830;
  --accent:#00d4ff; --green:#00ff9d; --yellow:#ffd166;
  --text:#c8d8e8; --sub:#4a6070; --danger:#ff4d6d;
}
* { box-sizing:border-box; margin:0; padding:0; }
body {
  background:var(--bg); color:var(--text);
  font-family:'IBM Plex Sans',sans-serif;
  min-height:100vh; display:flex; flex-direction:column;
  align-items:center; padding:40px 16px 60px;
}
body::before {
  content:''; position:fixed; inset:0; pointer-events:none; z-index:0;
  background-image:
    linear-gradient(rgba(0,212,255,.04) 1px, transparent 1px),
    linear-gradient(90deg, rgba(0,212,255,.04) 1px, transparent 1px);
  background-size:40px 40px;
}
.wrap { position:relative; z-index:1; width:100%; max-width:600px; }
header { text-align:center; margin-bottom:32px; }
.logo { font-size:11px; color:var(--accent); letter-spacing:4px;
        font-family:'IBM Plex Mono',monospace; margin-bottom:10px; }
h1 { font-size:24px; font-weight:600; color:#fff; }
h1 span { color:var(--accent); }
.sub-h { font-size:12px; color:var(--sub); margin-top:5px; }
.card { background:var(--card); border:1px solid var(--border);
        border-radius:6px; padding:20px; margin-bottom:12px; }
.slabel { font-family:'IBM Plex Mono',monospace; font-size:10px;
          color:var(--sub); letter-spacing:2px; margin-bottom:14px; }
.field + .field { margin-top:12px; }
label { display:block; font-size:11px; color:var(--sub); margin-bottom:5px; }
.file-input, .text-input {
  width:100%; background:#050709; border:1px solid var(--border); border-radius:4px;
  color:var(--text); font-family:'IBM Plex Mono',monospace; font-size:11px;
  padding:9px 12px; outline:none; cursor:pointer; transition:border-color .2s;
}
.file-input:hover, .text-input:hover, .text-input:focus { border-color:var(--accent); }
.row2 { display:grid; grid-template-columns:1fr 1fr; gap:10px; }
.hint { font-size:11px; color:var(--sub); margin-top:6px; line-height:1.5; }
.hint code { color:var(--accent); }
.fmt-toggle { display:flex; gap:8px; }
.fmt-btn {
  flex:1; background:#050709; border:1px solid var(--border); border-radius:4px;
  color:var(--sub); font-family:'IBM Plex Mono',monospace; font-size:12px;
  padding:10px; cursor:pointer; transition:all .2s; text-align:center;
  letter-spacing:1px; user-select:none;
}
.fmt-btn.active { border-color:var(--accent); color:var(--accent); background:rgba(0,212,255,.06); }
.stats { display:flex; gap:10px; margin-bottom:12px; }
.stat {
  flex:1; background:#050709; border:1px solid var(--border);
  border-radius:4px; padding:12px; text-align:center;
}
.stat-num { font-size:26px; font-weight:600; color:var(--accent);
            font-family:'IBM Plex Mono',monospace; }
.stat-num.g { color:var(--green); }
.stat-lbl { font-size:10px; color:var(--sub); margin-top:2px; }
.log {
  background:#020304; border:1px solid var(--border); border-radius:4px;
  font-family:'IBM Plex Mono',monospace; font-size:11px; color:var(--green);
  padding:12px; height:130px; overflow-y:auto; white-space:pre-wrap; word-break:break-all;
}
.log .err  { color:var(--danger); }
.log .warn { color:var(--yellow); }
.btn {
  width:100%; background:var(--accent); color:#000; border:none; border-radius:4px;
  font-family:'IBM Plex Mono',monospace; font-size:13px; font-weight:600;
  letter-spacing:2px; padding:15px; cursor:pointer; transition:opacity .2s, transform .1s;
}
.btn:hover  { opacity:.88; }
.btn:active { transform:scale(.99); }
.btn:disabled { opacity:.35; cursor:not-allowed; }
.banner {
  display:none; background:rgba(0,255,157,.08); border:1px solid var(--green);
  border-radius:4px; padding:13px; color:var(--green); font-size:13px;
  margin-bottom:12px; text-align:center;
}
footer { margin-top:36px; font-size:11px; color:var(--sub); text-align:center; }
</style>
</head>
<body>
<div class="wrap">

  <header>
    <div class="logo">PAJAK CONVERTER</div>
    <h1>XML <span>&rarr;</span> PajakExpress</h1>
    <p class="sub-h">Coretax DLP XML &rarr; ImporPajakKeluaran XLSX / CSV</p>
  </header>

  <div id="banner" class="banner"></div>

  <!-- 01 FILES -->
  <div class="card">
    <div class="slabel">01 / INPUT FILES</div>
    <div class="field">
      <label>File XML Coretax *</label>
      <input class="file-input" type="file" id="xml-file" accept=".xml">
    </div>
    <div class="field" id="tmpl-wrap">
      <label>Template XLSX (ImporPajakKeluaran_CSV_) &mdash; wajib untuk output XLSX</label>
      <input class="file-input" type="file" id="tmpl-file" accept=".xlsx">
    </div>
  </div>

  <!-- 02 NOMOR FAKTUR -->
  <div class="card">
    <div class="slabel">02 / NOMOR FAKTUR <span style="color:var(--danger)">*</span></div>
    <p class="hint" style="margin-bottom:12px;">
      NOMOR_FAKTUR tidak ada di XML Coretax &mdash; akan di-generate otomatis.<br>
      Default: <code>YYYYMM</code> + 5 digit urutan &nbsp;&rarr;&nbsp;
      <code id="nomor-ex">20260400001</code>
    </p>
    <div class="row2">
      <div class="field">
        <label>Prefix kustom (opsional)</label>
        <input class="text-input" type="text" id="prefix" placeholder="cth: INV atau kosongkan">
      </div>
      <div class="field">
        <label>Nomor urut awal</label>
        <input class="text-input" type="number" id="start-num" value="1" min="1">
      </div>
    </div>
  </div>

  <!-- 03 FORMAT -->
  <div class="card">
    <div class="slabel">03 / FORMAT OUTPUT</div>
    <div class="fmt-toggle">
      <div class="fmt-btn active" id="btn-xlsx">&#128202;&nbsp; XLSX</div>
      <div class="fmt-btn"        id="btn-csv"> &#128196;&nbsp; CSV</div>
    </div>
    <p class="hint" style="margin-top:10px;" id="fmt-hint">
      XLSX: mengisi template ImporPajakKeluaran (perlu upload template di atas).
    </p>
  </div>

  <!-- 04 PREVIEW -->
  <div class="card">
    <div class="slabel">04 / PREVIEW</div>
    <div class="stats">
      <div class="stat">
        <div class="stat-num"   id="inv-count">&#8212;</div>
        <div class="stat-lbl">Faktur (FK)</div>
      </div>
      <div class="stat">
        <div class="stat-num g" id="item-count">&#8212;</div>
        <div class="stat-lbl">Item (OF)</div>
      </div>
    </div>
    <div class="log" id="log">Pilih file XML untuk melihat preview&hellip;</div>
  </div>

  <button class="btn" id="conv-btn">&#9654;&nbsp; CONVERT &amp; DOWNLOAD</button>
  <footer>Coretax DLP XML &rarr; ImporPajakKeluaran &nbsp;|&nbsp; XLSX &amp; CSV</footer>
</div>

<script>
(function () {
  /* ── state ── */
  var outputFmt = 'xlsx';

  /* ── format toggle ── */
  document.getElementById('btn-xlsx').addEventListener('click', function () { setFmt('xlsx'); });
  document.getElementById('btn-csv').addEventListener('click',  function () { setFmt('csv');  });

  function setFmt(f) {
    outputFmt = f;
    document.getElementById('btn-xlsx').classList.toggle('active', f === 'xlsx');
    document.getElementById('btn-csv').classList.toggle('active',  f === 'csv');
    document.getElementById('tmpl-wrap').style.display = (f === 'xlsx') ? '' : 'none';
    document.getElementById('fmt-hint').textContent = (f === 'xlsx')
      ? 'XLSX: mengisi template ImporPajakKeluaran (perlu upload template di atas).'
      : 'CSV: UTF-8 BOM, flat file, langsung bisa diupload ke PajakExpress.';
  }

  /* ── nomor preview ── */
  function updateNomorPreview() {
    var prefix = document.getElementById('prefix').value.trim();
    var start  = parseInt(document.getElementById('start-num').value, 10) || 1;
    var num    = String(start).padStart(5, '0');
    var sample = prefix ? (prefix + num) : ('YYYYMM' + num);
    document.getElementById('nomor-ex').textContent = sample;
  }
  document.getElementById('prefix').addEventListener('input',  updateNomorPreview);
  document.getElementById('start-num').addEventListener('input', updateNomorPreview);
  updateNomorPreview();

  /* ── log helpers ── */
  function addLog(msg, cls) {
    var el   = document.getElementById('log');
    var span = document.createElement('span');
    if (cls) span.className = cls;
    span.textContent = msg + '\n';
    el.appendChild(span);
    el.scrollTop = el.scrollHeight;
  }
  function clearLog() { document.getElementById('log').innerHTML = ''; }

  /* ── XML preview (on file select) ── */
  document.getElementById('xml-file').addEventListener('change', function () {
    if (!this.files.length) return;
    var fd = new FormData();
    fd.append('xml', this.files[0]);
    clearLog();
    addLog('Membaca XML\u2026');

    fetch('/api/preview', { method: 'POST', body: fd })
      .then(function (res) { return res.json(); })
      .then(function (data) {
        if (data.error) { addLog('ERROR: ' + data.error, 'err'); return; }
        document.getElementById('inv-count').textContent  = data.invoices;
        document.getElementById('item-count').textContent = data.items;
        (data.warnings || []).forEach(function (w) { addLog('\u26a0 ' + w, 'warn'); });
        data.preview.forEach(function (p) { addLog(p); });
      })
      .catch(function (e) { addLog('Gagal: ' + e, 'err'); });
  });

  /* ── convert ── */
  document.getElementById('conv-btn').addEventListener('click', function () {
    var xmlFile  = document.getElementById('xml-file').files[0];
    var tmplFile = document.getElementById('tmpl-file').files[0];
    var prefix   = document.getElementById('prefix').value.trim();
    var startNum = parseInt(document.getElementById('start-num').value, 10) || 1;

    if (!xmlFile) { alert('Pilih file XML Coretax terlebih dahulu.'); return; }
    if (outputFmt === 'xlsx' && !tmplFile) {
      alert('Untuk output XLSX, upload dulu file template ImporPajakKeluaran_CSV_.xlsx.\nAtau ganti format ke CSV.');
      return;
    }

    var btn = document.getElementById('conv-btn');
    btn.disabled    = true;
    btn.textContent = 'SEDANG MENGONVERSI\u2026';
    clearLog();
    addLog('Memproses ' + xmlFile.name + '\u2026');

    var fd = new FormData();
    fd.append('xml',       xmlFile);
    fd.append('format',    outputFmt);
    fd.append('prefix',    prefix);
    fd.append('start_num', startNum);
    if (tmplFile) fd.append('template', tmplFile);

    fetch('/api/convert', { method: 'POST', body: fd })
      .then(function (res) {
        if (!res.ok) {
          return res.json().then(function (err) {
            throw new Error(err.error || res.statusText);
          });
        }
        return res.blob();
      })
      .then(function (blob) {
        /* ── FIX: append anchor to DOM sebelum click, revoke setelah delay ── */
        var ext  = (outputFmt === 'csv') ? '.csv' : '.xlsx';
        var name = xmlFile.name.replace(/\.xml$/i, '') + '_converted' + ext;
        var url  = URL.createObjectURL(blob);
        var a    = document.createElement('a');
        a.href        = url;
        a.download    = name;
        a.style.display = 'none';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(function () { URL.revokeObjectURL(url); }, 500);

        addLog('\u2713 Selesai \u2014 ' + name + ' sudah diunduh.');
        var banner = document.getElementById('banner');
        banner.textContent = '\u2713 Konversi berhasil! File: ' + name;
        banner.style.display = 'block';
      })
      .catch(function (e) {
        addLog('ERROR: ' + e.message, 'err');
      })
      .finally(function () {
        btn.disabled    = false;
        btn.textContent = '\u25b6  CONVERT & DOWNLOAD';
      });
  });

})();
</script>
</body>
</html>"""


# ─────────────────────────────────────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    # Gunakan Response() langsung — BUKAN render_template_string()
    # agar Jinja2 tidak memproses CSS/JS curly braces
    return Response(HTML, mimetype="text/html")


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
                f"[{i}] {inv['nama'] or '-'}  |  "
                f"{inv['tanggal_faktur'] or '-'}  |  "
                f"TrxCode={inv['kd_jenis_transaksi'] or '-'}  |  "
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
            buf     = write_csv(invoices, prefix, start)
            mime    = "text/csv"
            dl_name = "converted.csv"
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


# ─────────────────────────────────────────────────────────────────────────────
# Local development entry point
# Vercel memanggil `app` langsung tanpa blok ini.
# Jalankan lokal dengan: python index.py
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app.run(debug=True, port=5000)
