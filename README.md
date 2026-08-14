<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ZET BANK</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.1/jspdf.plugin.autotable.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"></script>
<style>
/* ============================================================
   ZET BANK — DESIGN TOKENS (identik dengan versi prototipe)
   ============================================================ */
:root{
  --zb-navy:#10233F;
  --zb-navy-2:#16304F;
  --zb-navy-3:#0B1930;
  --zb-paper:#EEF2EF;
  --zb-card:#FFFFFF;
  --zb-gold:#B8863B;
  --zb-gold-soft:#E7D5B0;
  --zb-ink:#182534;
  --zb-muted:#5B6B7A;
  --zb-line:#D7DDD9;
  --zb-green:#2E6F4E;
  --zb-green-soft:#E4F0E9;
  --zb-red:#B23A2E;
  --zb-red-soft:#F6E7E4;
  --zb-radius:10px;
  --zb-shadow:0 1px 2px rgba(16,35,63,.06), 0 8px 24px -12px rgba(16,35,63,.18);
  --font-display:'Fraunces', serif;
  --font-body:'Inter', system-ui, sans-serif;
  --font-mono:'IBM Plex Mono', monospace;
}
*{box-sizing:border-box;}
html,body{margin:0;padding:0;}
body{
  font-family:var(--font-body);
  background:var(--zb-paper);
  color:var(--zb-ink);
  min-height:100vh;
  -webkit-font-smoothing:antialiased;
}
button{font-family:inherit;cursor:pointer;}
input,select{font-family:inherit;}
::selection{background:var(--zb-gold-soft);}
:focus-visible{outline:2px solid var(--zb-gold);outline-offset:2px;}
.hidden{display:none !important;}
.mono{font-family:var(--font-mono);font-variant-numeric:tabular-nums;}
.stamp{
  display:inline-flex;align-items:center;gap:6px;
  font-family:var(--font-mono);font-size:12px;letter-spacing:.04em;
  border:1px solid var(--zb-gold);color:var(--zb-navy);
  background:var(--zb-gold-soft);
  padding:3px 9px;border-radius:999px;
  transform:rotate(-1deg);
}
.auth-shell{
  min-height:100vh;display:flex;align-items:center;justify-content:center;
  padding:32px 16px;
  background:
    radial-gradient(1100px 500px at 15% -10%, #1A3A62 0%, transparent 60%),
    var(--zb-navy);
}
.passbook{
  width:100%;max-width:960px;
  display:grid;grid-template-columns:1.1fr 1fr;
  background:var(--zb-card);
  border-radius:16px;
  box-shadow:0 40px 80px -30px rgba(0,0,0,.55);
  overflow:hidden;
  min-height:560px;
}
@media (max-width:820px){ .passbook{grid-template-columns:1fr;} .passbook-cover{display:none;} }
.passbook-cover{
  background:linear-gradient(180deg, var(--zb-navy-2), var(--zb-navy-3));
  color:#fff;position:relative;padding:44px 40px;
  display:flex;flex-direction:column;justify-content:space-between;
}
.passbook-cover::before{
  content:"";position:absolute;left:0;top:0;bottom:0;width:14px;
  background:repeating-linear-gradient(to bottom, var(--zb-gold) 0 3px, transparent 3px 14px);
  opacity:.55;
}
.brand-mark{display:flex;align-items:center;gap:12px;}
.brand-badge{
  width:46px;height:46px;border-radius:10px;
  background:linear-gradient(145deg, var(--zb-gold), #8f6423);
  display:flex;align-items:center;justify-content:center;
  font-family:var(--font-display);font-weight:700;color:var(--zb-navy-3);font-size:20px;
  box-shadow:inset 0 1px 0 rgba(255,255,255,.4);
}
.brand-name{font-family:var(--font-display);font-size:26px;font-weight:600;letter-spacing:.02em;}
.brand-tag{font-size:12.5px;color:#B9C7DA;margin-top:2px;}
.cover-quote{font-family:var(--font-display);font-size:22px;line-height:1.4;color:#EAF0F7;max-width:36ch;}
.cover-meta{font-size:12px;color:#8FA3BE;font-family:var(--font-mono);}
.auth-panel{padding:40px 40px;display:flex;flex-direction:column;}
.auth-tabs{display:flex;gap:4px;background:var(--zb-paper);padding:4px;border-radius:10px;margin-bottom:26px;}
.auth-tab{flex:1;border:none;background:transparent;padding:10px 8px;border-radius:8px;font-size:13.5px;font-weight:600;color:var(--zb-muted);}
.auth-tab.active{background:var(--zb-card);color:var(--zb-navy);box-shadow:var(--zb-shadow);}
.field{margin-bottom:14px;}
.field label{display:block;font-size:12.5px;font-weight:600;color:var(--zb-muted);margin-bottom:6px;}
.field input,.field select{width:100%;padding:11px 12px;border:1px solid var(--zb-line);border-radius:8px;font-size:14.5px;background:#fff;color:var(--zb-ink);}
.field input:focus,.field select:focus{border-color:var(--zb-gold);}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
.hint{font-size:12px;color:var(--zb-muted);margin-top:10px;line-height:1.5;}
.btn{border:none;border-radius:8px;padding:11px 18px;font-weight:600;font-size:14px;display:inline-flex;align-items:center;justify-content:center;gap:7px;transition:transform .05s ease;}
.btn:active{transform:translateY(1px);}
.btn:disabled{opacity:.5;cursor:not-allowed;}
.btn-primary{background:var(--zb-navy);color:#fff;width:100%;}
.btn-primary:hover{background:var(--zb-navy-2);}
.btn-gold{background:var(--zb-gold);color:#fff;}
.btn-ghost{background:transparent;color:var(--zb-navy);border:1px solid var(--zb-line);}
.btn-ghost:hover{border-color:var(--zb-navy);}
.btn-sm{padding:7px 12px;font-size:12.5px;border-radius:7px;}
.btn-block{width:100%;}
.link-btn{background:none;border:none;color:var(--zb-navy);font-weight:600;font-size:13px;text-decoration:underline;padding:0;}
.divider{display:flex;align-items:center;gap:10px;color:var(--zb-muted);font-size:12px;margin:18px 0;}
.divider::before,.divider::after{content:"";flex:1;height:1px;background:var(--zb-line);}
.error-msg{background:var(--zb-red-soft);color:var(--zb-red);font-size:12.5px;padding:9px 12px;border-radius:7px;margin-bottom:12px;}
.ok-msg{background:var(--zb-green-soft);color:var(--zb-green);font-size:12.5px;padding:9px 12px;border-radius:7px;margin-bottom:12px;}
.app-shell{min-height:100vh;display:none;flex-direction:column;}
.app-shell.active{display:flex;}
.app-header{background:var(--zb-navy);color:#fff;padding:14px 24px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;position:sticky;top:0;z-index:20;}
.app-header .brand-mark .brand-name{font-size:19px;color:#fff;}
.app-header .brand-mark .brand-badge{width:36px;height:36px;font-size:16px;}
.header-loc{font-size:11.5px;color:#9FB3CC;font-family:var(--font-mono);margin-top:1px;}
.header-left{display:flex;align-items:center;gap:18px;}
.saldo-total{background:rgba(184,134,59,.16);border:1px solid rgba(184,134,59,.5);border-radius:9px;padding:6px 14px;}
.saldo-total .lbl{font-size:10px;text-transform:uppercase;letter-spacing:.04em;color:#E7D5B0;}
.saldo-total .val{font-family:var(--font-mono);font-size:16px;font-weight:600;color:#fff;margin-top:1px;}
.header-right{display:flex;align-items:center;gap:14px;}
.clock{font-family:var(--font-mono);font-size:13px;color:#CFE0F2;letter-spacing:.03em;}
.who{font-size:12.5px;color:#CFE0F2;text-align:right;}
.who b{color:#fff;}
.icon-btn{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.16);color:#fff;border-radius:7px;padding:7px 10px;font-size:12px;display:inline-flex;gap:6px;align-items:center;}
.icon-btn:hover{background:rgba(255,255,255,.16);}
.toolbar{background:#fff;border-bottom:1px solid var(--zb-line);padding:12px 24px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;}
.toolbar .spacer{flex:1;}
.tbtn{background:var(--zb-paper);border:1px solid var(--zb-line);color:var(--zb-navy);border-radius:8px;padding:8px 12px;font-size:12.5px;font-weight:600;display:inline-flex;align-items:center;gap:6px;position:relative;}
.tbtn:hover{border-color:var(--zb-navy);}
.tbtn.primary{background:var(--zb-navy);color:#fff;border-color:var(--zb-navy);}
.tbtn.primary:hover{background:var(--zb-navy-2);}
.badge{background:var(--zb-red);color:#fff;font-size:10px;font-weight:700;border-radius:999px;min-width:16px;height:16px;display:inline-flex;align-items:center;justify-content:center;padding:0 4px;}
.privacy-note{margin:14px 24px 0;background:var(--zb-gold-soft);border:1px solid var(--zb-gold);color:#5A4520;font-size:12.5px;padding:9px 14px;border-radius:8px;}
.filters{margin:14px 24px 0;background:#fff;border:1px solid var(--zb-line);border-radius:10px;padding:14px 16px;display:flex;gap:12px;flex-wrap:wrap;align-items:end;}
.filters .field{margin-bottom:0;min-width:150px;flex:1;}
.filters .field input,.filters .field select{padding:9px 10px;font-size:13px;}
.content{padding:16px 24px 40px;flex:1;}
.ledger-card{background:#fff;border:1px solid var(--zb-line);border-radius:12px;overflow:hidden;box-shadow:var(--zb-shadow);}
.ledger-head{padding:16px 18px 8px;display:flex;justify-content:space-between;align-items:baseline;}
.ledger-head h2{font-family:var(--font-display);font-size:19px;margin:0;}
.ledger-count{font-size:12px;color:var(--zb-muted);font-family:var(--font-mono);}
table{width:100%;border-collapse:collapse;font-size:13px;}
thead th{text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.05em;color:var(--zb-muted);padding:10px 18px;border-bottom:1px solid var(--zb-line);background:var(--zb-paper);}
tbody td{padding:11px 18px;border-bottom:1px solid #EEF0EE;vertical-align:middle;}
tbody tr:hover{background:#FAFBF9;}
.acct{font-family:var(--font-mono);font-size:12px;color:var(--zb-navy);background:var(--zb-paper);padding:2px 7px;border-radius:5px;}
.tag{font-size:11px;font-weight:700;padding:3px 9px;border-radius:999px;display:inline-block;}
.tag-setor{background:var(--zb-green-soft);color:var(--zb-green);}
.tag-tarik{background:var(--zb-red-soft);color:var(--zb-red);}
.amt-in{color:var(--zb-green);font-family:var(--font-mono);font-weight:600;}
.amt-out{color:var(--zb-red);font-family:var(--font-mono);font-weight:600;}
.amt-hidden{color:var(--zb-muted);font-family:var(--font-mono);}
.empty-state{padding:60px 20px;text-align:center;color:var(--zb-muted);}
.empty-state .big{font-family:var(--font-display);font-size:20px;color:var(--zb-navy);margin-bottom:6px;}
.footer-credit{text-align:center;color:var(--zb-muted);font-size:11.5px;padding:22px;}
.modal-backdrop{position:fixed;inset:0;background:rgba(11,25,48,.55);display:none;align-items:flex-start;justify-content:center;z-index:100;padding:40px 16px;overflow-y:auto;}
.modal-backdrop.active{display:flex;}
.modal{background:#fff;border-radius:14px;width:100%;max-width:560px;box-shadow:0 40px 80px -20px rgba(0,0,0,.4);animation:pop .16s ease;}
.modal.wide{max-width:760px;}
@keyframes pop{from{opacity:0;transform:translateY(8px) scale(.98);}to{opacity:1;transform:none;}}
.modal-head{padding:18px 22px;border-bottom:1px solid var(--zb-line);display:flex;justify-content:space-between;align-items:center;}
.modal-head h3{margin:0;font-family:var(--font-display);font-size:19px;}
.modal-close{background:none;border:none;font-size:18px;color:var(--zb-muted);width:30px;height:30px;border-radius:7px;}
.modal-close:hover{background:var(--zb-paper);}
.modal-body{padding:20px 22px;max-height:70vh;overflow-y:auto;}
.modal-foot{padding:16px 22px;border-top:1px solid var(--zb-line);display:flex;justify-content:flex-end;gap:10px;}
.subtabs{display:flex;gap:6px;margin-bottom:16px;border-bottom:1px solid var(--zb-line);}
.subtab{background:none;border:none;padding:9px 4px;margin-right:14px;font-size:13px;font-weight:600;color:var(--zb-muted);border-bottom:2px solid transparent;}
.subtab.active{color:var(--zb-navy);border-color:var(--zb-gold);}
.mini-table{width:100%;border-collapse:collapse;font-size:12.5px;margin-top:10px;}
.mini-table th{text-align:left;color:var(--zb-muted);font-size:10.5px;text-transform:uppercase;padding:6px 8px;border-bottom:1px solid var(--zb-line);}
.mini-table td{padding:8px;border-bottom:1px solid #F0F1EE;}
.autocomplete-list{border:1px solid var(--zb-line);border-radius:8px;margin-top:4px;max-height:180px;overflow-y:auto;background:#fff;position:relative;z-index:5;}
.ac-item{padding:9px 12px;font-size:13px;cursor:pointer;}
.ac-item:hover{background:var(--zb-paper);}
.total-box{background:var(--zb-navy);color:#fff;border-radius:10px;padding:14px 16px;display:flex;justify-content:space-between;align-items:center;margin-top:4px;}
.total-box .lbl{font-size:12px;color:#B9C7DA;}
.total-box .val{font-family:var(--font-mono);font-size:20px;font-weight:600;}
.saldo-card{background:var(--zb-paper);border:1px dashed var(--zb-gold);border-radius:10px;padding:16px;margin-bottom:14px;}
.saldo-card .lbl{font-size:11.5px;color:var(--zb-muted);text-transform:uppercase;letter-spacing:.04em;}
.saldo-card .val{font-family:var(--font-mono);font-size:26px;font-weight:600;color:var(--zb-navy);margin-top:4px;}
.spinner-overlay{position:fixed;inset:0;background:rgba(238,242,239,.6);display:none;align-items:center;justify-content:center;z-index:200;font-family:var(--font-mono);color:var(--zb-navy);font-size:13px;}
.spinner-overlay.active{display:flex;}
</style>
</head>
<body>

<div id="spinner" class="spinner-overlay">⏳ Memproses…</div>

<!-- ============================================================
     AUTH SCREENS
     ============================================================ -->
<div id="screen-auth" class="auth-shell">
  <div class="passbook">
    <div class="passbook-cover">
      <div class="brand-mark">
        <div class="brand-badge">Z</div>
        <div><div class="brand-name">ZET BANK</div><div class="brand-tag">Buku Tabungan Digital Warga</div></div>
      </div>
      <div><p class="cover-quote">"Setiap setoran tercatat, setiap saldo terjaga."</p></div>
      <div class="cover-meta" id="cover-clock">--:--:--</div>
    </div>
    <div class="auth-panel">
      <div id="pane-login">
        <div class="auth-tabs">
          <button class="auth-tab" data-authtab="staf" onclick="switchAuthTab('staf')">Login Staf</button>
          <button class="auth-tab active" data-authtab="nasabah" onclick="switchAuthTab('nasabah')">Nasabah</button>
        </div>
        <div id="login-error" class="error-msg hidden"></div>
        <div id="authtab-staf" class="hidden">
          <div class="field"><label>Email</label><input type="email" id="staf-email" placeholder="nama@zetbank.id"></div>
          <div class="field"><label>Password</label><input type="password" id="staf-password" placeholder="••••••••"></div>
          <button class="btn btn-primary" onclick="doLoginStaf()">Masuk</button>
        </div>
        <div id="authtab-nasabah">
          <div class="field"><label>No. Rekening / Email</label><input type="text" id="nas-id" placeholder="No. rekening atau email"></div>
          <div class="field"><label>Password</label><input type="password" id="nas-password" placeholder="••••••••"></div>
          <button class="btn btn-primary" onclick="doLoginNasabah()">Masuk</button>
          <div class="divider">atau</div>
          <button class="btn btn-ghost btn-block" onclick="showRegister()">Daftar Akun Baru</button>
        </div>
      </div>
      <div id="pane-register" class="hidden">
        <button class="link-btn" onclick="backToLogin()" style="margin-bottom:14px;">← Kembali ke Login</button>
        <h3 style="font-family:var(--font-display);margin:0 0 4px;">Daftar Akun Baru</h3>
        <p class="hint" style="margin-top:0;">Pendaftaran akan diverifikasi oleh Admin sebelum akun aktif dan tercatat sebagai nasabah resmi.</p>
        <div id="register-error" class="error-msg hidden"></div>
        <div id="register-ok" class="ok-msg hidden"></div>
        <div class="field"><label>Nama Lengkap</label><input type="text" id="reg-nama"></div>
        <div class="row2">
          <div class="field"><label>RT (opsional)</label><input type="text" id="reg-rt"></div>
          <div class="field"><label>RW (opsional)</label><input type="text" id="reg-rw"></div>
        </div>
        <div class="field"><label>Alamat Rumah *wajib</label><input type="text" id="reg-alamat"></div>
        <div class="field"><label>No. HP / Email (opsional)</label><input type="text" id="reg-kontak"></div>
        <div class="field"><label>Buat Password</label><input type="password" id="reg-password" placeholder="Minimal 6 karakter"></div>
        <button class="btn btn-primary" onclick="submitRegistration()">Kirim Pendaftaran</button>
      </div>
    </div>
  </div>
</div>

<div id="screen-forcepw" class="auth-shell hidden">
  <div class="passbook" style="grid-template-columns:1fr;max-width:440px;">
    <div class="auth-panel">
      <div class="brand-mark" style="margin-bottom:18px;">
        <div class="brand-badge" style="background:linear-gradient(145deg,var(--zb-gold),#8f6423);">Z</div>
        <div><div class="brand-name" style="color:var(--zb-navy);">ZET BANK</div></div>
      </div>
      <h3 style="font-family:var(--font-display);margin:0 0 4px;">Wajib Ganti Password</h3>
      <p class="hint" style="margin-top:0;">Ini pertama kalinya Anda login. Demi keamanan, buat password baru sebelum melanjutkan ke aplikasi.</p>
      <div id="forcepw-error" class="error-msg hidden"></div>
      <div class="field"><label>Password Baru</label><input type="password" id="forcepw-1"></div>
      <div class="field"><label>Ulangi Password Baru</label><input type="password" id="forcepw-2"></div>
      <button class="btn btn-primary" onclick="submitForcePw()">Simpan &amp; Lanjutkan</button>
    </div>
  </div>
</div>

<!-- ============================================================
     MAIN APP SHELL
     ============================================================ -->
<div id="app" class="app-shell">
  <header class="app-header">
    <div class="header-left">
      <div class="brand-mark">
        <div class="brand-badge">Z</div>
        <div><div class="brand-name">ZET BANK</div><div class="header-loc">Bank Keuangan Warga</div></div>
      </div>
      <div class="saldo-total hidden" id="saldo-total-box">
        <div class="lbl">Saldo Total Nasabah</div>
        <div class="val" id="saldo-total-val">Rp 0</div>
      </div>
    </div>
    <div class="header-right">
      <div class="clock" id="app-clock">--:--:--</div>
      <div class="who">Masuk sebagai: <b id="who-name">-</b><br><span id="who-role" style="font-family:var(--font-mono);font-size:10.5px;"></span></div>
      <button class="icon-btn" onclick="openModal('modal-changepw')">🔑 Ganti Password</button>
      <button class="icon-btn" onclick="doLogout()">⏻ Keluar</button>
    </div>
  </header>

  <div class="toolbar" id="toolbar"></div>
  <div class="privacy-note hidden" id="privacy-note">
    🔒 Rincian nominal transaksi bersifat pribadi — hanya dapat dilihat oleh Anda sendiri, serta Admin dan Manajer (Super Admin). Nasabah lain tidak dapat melihat transaksi Anda.
  </div>

  <div class="filters">
    <div class="field"><label>Dari Tanggal</label><input type="date" id="f-from" onchange="renderLedger()"></div>
    <div class="field"><label>Sampai Tanggal</label><input type="date" id="f-to" onchange="renderLedger()"></div>
    <div class="field"><label>Cari Nasabah</label><input type="text" id="f-search" placeholder="Nama / no. rekening" oninput="renderLedger()"></div>
    <div class="field"><label>Tipe</label>
      <select id="f-tipe" onchange="renderLedger()"><option value="">Semua Tipe</option><option value="setor">Setor</option><option value="tarik">Tarik</option></select>
    </div>
    <button class="btn btn-ghost btn-sm" onclick="resetFilters()">Reset Filter</button>
  </div>

  <div class="content">
    <div class="ledger-card">
      <div class="ledger-head"><h2>Buku Besar Transaksi</h2><span class="ledger-count" id="ledger-count">0 catatan</span></div>
      <div style="overflow-x:auto;">
        <table>
          <thead><tr>
            <th>Tanggal / Jam</th><th>Nasabah</th><th>No. Rek</th><th>Tipe</th><th>Nominal</th><th>Saldo</th><th>Keterangan</th><th></th>
          </tr></thead>
          <tbody id="ledger-body"></tbody>
        </table>
      </div>
      <div id="ledger-empty" class="empty-state hidden"><div class="big">Belum ada transaksi</div><div>Transaksi yang tercatat akan muncul di sini.</div></div>
    </div>
    <div class="footer-credit">ZET BANK — Manajemen Keuangan Warga · Tersambung ke database Supabase (data tersimpan aman di cloud)</div>
  </div>
</div>

<!-- ============================================================
     MODAL: TAMBAH TRANSAKSI
     ============================================================ -->
<div class="modal-backdrop" id="modal-transaksi">
  <div class="modal">
    <div class="modal-head"><h3>＋ Tambah Transaksi</h3><button class="modal-close" onclick="closeModal('modal-transaksi')">✕</button></div>
    <div class="modal-body">
      <div id="trx-error" class="error-msg hidden"></div>
      <div id="trx-ok" class="ok-msg hidden">✓ Transaksi tersimpan.</div>
      <div class="field">
        <label>Pemilik Rekening (Nasabah)</label>
        <input type="text" id="trx-nasabah-search" placeholder="Ketik untuk mencari nasabah" onclick="renderNasabahAC('')" oninput="renderNasabahAC(this.value)">
        <div id="trx-nasabah-ac" class="autocomplete-list hidden"></div>
        <input type="hidden" id="trx-nasabah-id">
      </div>
      <div class="row2">
        <div class="field"><label>Tanggal</label><input type="date" id="trx-tanggal"></div>
        <div class="field"><label>Jam</label><input type="time" id="trx-jam"></div>
      </div>
      <div class="field">
        <label>Tipe Transaksi</label>
        <div style="display:flex;gap:10px;">
          <button type="button" class="btn btn-ghost" id="trx-tipe-setor" style="flex:1;" onclick="setTrxTipe('setor')">Setor Tunai</button>
          <button type="button" class="btn btn-ghost" id="trx-tipe-tarik" style="flex:1;" onclick="setTrxTipe('tarik')">Tarik Saldo</button>
        </div>
        <input type="hidden" id="trx-tipe" value="setor">
      </div>
      <div class="field">
        <label>Nominal (Rp)</label>
        <select id="trx-nominal-preset" onchange="applyNominalPreset()" style="margin-bottom:8px;"><option value="">Pilih nominal cepat… (opsional)</option></select>
        <input type="text" id="trx-nominal" placeholder="Masukkan nominal, mis. 50000" oninput="updateTrxTotal()">
      </div>
      <div class="total-box"><span class="lbl">Total Nominal</span><span class="val" id="trx-total">Rp 0</span></div>
      <div class="field" style="margin-top:14px;"><label>Keterangan (opsional)</label><input type="text" id="trx-keterangan"></div>
      <p class="hint">Tekan OK untuk merekam transaksi ini &amp; langsung tambah transaksi baru untuk nasabah yang sama.</p>
    </div>
    <div class="modal-foot">
      <button class="btn btn-ghost" onclick="closeModal('modal-transaksi')">Batal</button>
      <button class="btn btn-ghost" onclick="cetakStruk()">🖨️ Cetak Struk</button>
      <button class="btn btn-gold" onclick="simpanTransaksi(true)">OK</button>
      <button class="btn btn-primary" onclick="simpanTransaksi(false)">Selesai &amp; Tutup</button>
    </div>
  </div>
</div>

<!-- ============================================================
     MODAL: EDIT TRANSAKSI (khusus Manajer)
     ============================================================ -->
<div class="modal-backdrop" id="modal-edit-transaksi">
  <div class="modal">
    <div class="modal-head"><h3>✎ Edit Transaksi</h3><button class="modal-close" onclick="closeModal('modal-edit-transaksi')">✕</button></div>
    <div class="modal-body">
      <div id="edit-trx-error" class="error-msg hidden"></div>
      <p class="hint" style="margin-top:0;" id="edit-trx-nasabah-info"></p>
      <input type="hidden" id="edit-trx-id">
      <div class="row2">
        <div class="field"><label>Tanggal</label><input type="date" id="edit-trx-tanggal"></div>
        <div class="field"><label>Jam</label><input type="time" id="edit-trx-jam"></div>
      </div>
      <div class="field">
        <label>Tipe Transaksi</label>
        <div style="display:flex;gap:10px;">
          <button type="button" class="btn btn-ghost" id="edit-trx-tipe-setor" style="flex:1;" onclick="setEditTrxTipe('setor')">Setor Tunai</button>
          <button type="button" class="btn btn-ghost" id="edit-trx-tipe-tarik" style="flex:1;" onclick="setEditTrxTipe('tarik')">Tarik Saldo</button>
        </div>
        <input type="hidden" id="edit-trx-tipe" value="setor">
      </div>
      <div class="field"><label>Nominal (Rp)</label><input type="text" id="edit-trx-nominal"></div>
      <div class="field"><label>Keterangan (opsional)</label><input type="text" id="edit-trx-keterangan"></div>
      <p class="hint">Perubahan ini hanya dapat dilakukan oleh Manajer (Super Admin).</p>
    </div>
    <div class="modal-foot">
      <button class="btn btn-ghost" onclick="closeModal('modal-edit-transaksi')">Batal</button>
      <button class="btn btn-primary" onclick="simpanEditTransaksi()">Simpan Perubahan</button>
    </div>
  </div>
</div>

<!-- ============================================================
     MODAL: KELOLA NOMINAL UANG
     ============================================================ -->
<div class="modal-backdrop" id="modal-nominal">
  <div class="modal">
    <div class="modal-head"><h3>⚙ Kelola Nominal Uang</h3><button class="modal-close" onclick="closeModal('modal-nominal')">✕</button></div>
    <div class="modal-body">
      <p class="hint" style="margin-top:0;">Daftar nominal cepat saat mencatat transaksi. Hanya Manajer dan Admin yang dapat mengubah.</p>
      <table class="mini-table"><thead><tr><th>Label</th><th>Nominal (Rp)</th><th></th></tr></thead><tbody id="nominal-body"></tbody></table>
      <div class="row2" style="margin-top:16px;">
        <div class="field"><label>Label</label><input type="text" id="new-nominal-label" placeholder="mis. Setoran Rutin"></div>
        <div class="field"><label>Nominal (Rp)</label><input type="text" id="new-nominal-value" placeholder="50000"></div>
      </div>
      <button class="btn btn-gold btn-sm" onclick="tambahNominal()">Tambah</button>
    </div>
    <div class="modal-foot"><button class="btn btn-primary" onclick="closeModal('modal-nominal')">Tutup</button></div>
  </div>
</div>

<!-- ============================================================
     MODAL: BUKU TABUNGAN NASABAH
     ============================================================ -->
<div class="modal-backdrop" id="modal-bukutabungan">
  <div class="modal wide">
    <div class="modal-head"><h3>🔍 Buku Tabungan Nasabah</h3><button class="modal-close" onclick="closeModal('modal-bukutabungan')">✕</button></div>
    <div class="modal-body">
      <div class="field">
        <label>Cari Nasabah</label>
        <input type="text" id="bt-search" placeholder="Nama atau no. rekening" oninput="renderBukuTabunganAC(this.value)" onclick="renderBukuTabunganAC(this.value)">
        <div id="bt-ac" class="autocomplete-list hidden"></div>
      </div>
      <div id="bt-detail"><p class="hint">Pilih nasabah untuk menampilkan riwayat dan saldo buku tabungannya.</p></div>
    </div>
    <div class="modal-foot">
      <button class="btn btn-ghost" id="bt-download" onclick="unduhBukuTabunganPDF()" disabled>📄 Unduh Buku Tabungan</button>
      <button class="btn btn-primary" onclick="closeModal('modal-bukutabungan')">Tutup</button>
    </div>
  </div>
</div>

<!-- ============================================================
     MODAL: KELOLA PENGGUNA
     ============================================================ -->
<div class="modal-backdrop" id="modal-pengguna">
  <div class="modal wide">
    <div class="modal-head"><h3>👷 Kelola Pengguna</h3><button class="modal-close" onclick="closeModal('modal-pengguna')">✕</button></div>
    <div class="modal-body">
      <div class="subtabs">
        <button class="subtab active" data-usertab="nasabah" onclick="switchUserTab('nasabah')">Nasabah</button>
        <button class="subtab" data-usertab="admin" onclick="switchUserTab('admin')">Akun Admin</button>
      </div>
      <div id="usertab-nasabah">
        <p class="hint" style="margin-top:0;">Input oleh Super Admin berdasarkan surat kuasa nasabah — langsung aktif tanpa antre verifikasi.</p>
        <table class="mini-table"><thead><tr><th>Nama</th><th>No. Rek</th><th>RT/RW</th><th>Alamat</th><th></th></tr></thead><tbody id="nasabah-body"></tbody></table>
        <div class="row2" style="margin-top:16px;"><div class="field"><label>Nama Nasabah Baru</label><input type="text" id="new-nasabah-nama"></div></div>
        <div class="row2">
          <div class="field"><label>RT (opsional)</label><input type="text" id="new-nasabah-rt"></div>
          <div class="field"><label>RW (opsional)</label><input type="text" id="new-nasabah-rw"></div>
        </div>
        <div class="field"><label>Alamat Rumah *wajib</label><input type="text" id="new-nasabah-alamat"></div>
        <button class="btn btn-gold btn-sm" onclick="tambahNasabahLangsung()">Tambah</button>
        <p class="hint">Password awal dibuat otomatis: <b>123456</b>. Berlaku sekali — nasabah wajib menggantinya saat login pertama.</p>
      </div>
      <div id="usertab-admin" class="hidden">
        <table class="mini-table"><thead><tr><th>Nama</th><th>Email</th><th>Peran</th><th></th></tr></thead><tbody id="admin-body"></tbody></table>
        <div class="row2" style="margin-top:16px;">
          <div class="field"><label>Nama Admin Baru</label><input type="text" id="new-admin-nama"></div>
          <div class="field"><label>Email</label><input type="email" id="new-admin-email"></div>
        </div>
        <button class="btn btn-gold btn-sm" onclick="tambahAdmin()">Tambah</button>
      </div>
    </div>
    <div class="modal-foot"><button class="btn btn-primary" onclick="closeModal('modal-pengguna')">Tutup</button></div>
  </div>
</div>

<!-- ============================================================
     MODAL: EDIT NASABAH (khusus Manajer)
     ============================================================ -->
<div class="modal-backdrop" id="modal-edit-nasabah">
  <div class="modal">
    <div class="modal-head"><h3>✎ Edit Data Nasabah</h3><button class="modal-close" onclick="closeModal('modal-edit-nasabah')">✕</button></div>
    <div class="modal-body">
      <div id="edit-nasabah-error" class="error-msg hidden"></div>
      <input type="hidden" id="edit-nasabah-id">
      <div class="field"><label>No. Rekening (tidak dapat diubah)</label><input type="text" id="edit-nasabah-norek" disabled style="background:var(--zb-paper);color:var(--zb-muted);"></div>
      <div class="field"><label>Nama</label><input type="text" id="edit-nasabah-nama"></div>
      <div class="row2">
        <div class="field"><label>RT (opsional)</label><input type="text" id="edit-nasabah-rt"></div>
        <div class="field"><label>RW (opsional)</label><input type="text" id="edit-nasabah-rw"></div>
      </div>
      <div class="field"><label>Alamat Rumah *wajib</label><input type="text" id="edit-nasabah-alamat"></div>
      <div class="field"><label>No. HP / Email (opsional)</label><input type="text" id="edit-nasabah-kontak"></div>
      <p class="hint">No. Rekening dan Password tidak dapat diubah lewat form ini.</p>
    </div>
    <div class="modal-foot">
      <button class="btn btn-ghost" onclick="closeModal('modal-edit-nasabah')">Batal</button>
      <button class="btn btn-primary" onclick="simpanEditNasabah()">Simpan Perubahan</button>
    </div>
  </div>
</div>

<!-- ============================================================
     MODAL: VERIFIKASI PENDAFTARAN
     ============================================================ -->
<div class="modal-backdrop" id="modal-verifikasi">
  <div class="modal wide">
    <div class="modal-head"><h3>🧾 Verifikasi Pendaftaran Nasabah</h3><button class="modal-close" onclick="closeModal('modal-verifikasi')">✕</button></div>
    <div class="modal-body">
      <table class="mini-table"><thead><tr><th>Nama</th><th>RT/RW</th><th>Alamat</th><th>Kontak</th><th></th></tr></thead><tbody id="verifikasi-body"></tbody></table>
      <div id="verifikasi-empty" class="empty-state hidden" style="padding:30px;">Tidak ada pendaftaran menunggu verifikasi.</div>
    </div>
    <div class="modal-foot"><button class="btn btn-primary" onclick="closeModal('modal-verifikasi')">Tutup</button></div>
  </div>
</div>

<!-- ============================================================
     MODAL: GANTI PASSWORD
     ============================================================ -->
<div class="modal-backdrop" id="modal-changepw">
  <div class="modal">
    <div class="modal-head"><h3>🔑 Ganti Password</h3><button class="modal-close" onclick="closeModal('modal-changepw')">✕</button></div>
    <div class="modal-body">
      <div id="changepw-error" class="error-msg hidden"></div>
      <div class="field"><label>Password Baru</label><input type="password" id="cpw-1"></div>
      <div class="field"><label>Ulangi Password Baru</label><input type="password" id="cpw-2"></div>
    </div>
    <div class="modal-foot">
      <button class="btn btn-ghost" onclick="closeModal('modal-changepw')">Batal</button>
      <button class="btn btn-primary" onclick="submitChangePw()">Simpan</button>
    </div>
  </div>
</div>

<script>
/* ================================================================
   ZET BANK — KONEKSI SUPABASE (database & autentikasi nyata)
   ================================================================ */
const SUPABASE_URL = 'https://rodgtzzayzywoxzengnt.supabase.co';
const SUPABASE_ANON_KEY = 'sb_publishable_pFBWIwSUibtC6tUSYl4Qbg_90qRElFB';
const EDGE_FUNCTION_URL = SUPABASE_URL + '/functions/v1/admin-actions';
const sb = supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

let currentProfile = null;   // baris profil pengguna yang sedang login
let pendingForceUser = null; // dipakai saat alur wajib-ganti-password

function rupiah(n){ return 'Rp '+ (Number(n)||0).toLocaleString('id-ID'); }
function todayStr(){ return new Date().toISOString().slice(0,10); }
function nowTimeStr(){ return new Date().toTimeString().slice(0,5); }
function formatTgl(iso){ if(!iso) return '-'; const [y,m,d]=iso.split('-'); return `${d}/${m}/${y}`; }
function showSpinner(on){ document.getElementById('spinner').classList.toggle('active', !!on); }
function hideErr(id){ document.getElementById(id).classList.add('hidden'); }
function showErr(id, msg){ const el=document.getElementById(id); el.textContent = msg; el.classList.remove('hidden'); }
async function edgeCall(action, payload){
  const { data:{session} } = await sb.auth.getSession();
  const res = await fetch(EDGE_FUNCTION_URL, {
    method:'POST',
    headers:{ 'Content-Type':'application/json', 'Authorization':'Bearer '+session.access_token },
    body: JSON.stringify({ action, ...payload })
  });
  return res.json();
}

/* ---------- Clock ---------- */
function tickClock(){
  const t = new Date().toLocaleTimeString('id-ID');
  const c1=document.getElementById('cover-clock'); if(c1) c1.textContent=t;
  const c2=document.getElementById('app-clock'); if(c2) c2.textContent=t;
}
setInterval(tickClock,1000); tickClock();

/* ================================================================
   AUTH TAB / SCREEN SWITCHING
   ================================================================ */
function switchAuthTab(tab){
  document.querySelectorAll('.auth-tab').forEach(b=>b.classList.remove('active'));
  document.querySelector(`[data-authtab="${tab}"]`).classList.add('active');
  document.getElementById('authtab-staf').classList.toggle('hidden', tab!=='staf');
  document.getElementById('authtab-nasabah').classList.toggle('hidden', tab!=='nasabah');
  hideErr('login-error');
}
function showRegister(){ document.getElementById('pane-login').classList.add('hidden'); document.getElementById('pane-register').classList.remove('hidden'); }
function backToLogin(){ document.getElementById('pane-register').classList.add('hidden'); document.getElementById('pane-login').classList.remove('hidden'); }

/* ================================================================
   LOGIN — STAF
   ================================================================ */
async function doLoginStaf(){
  hideErr('login-error');
  const email = document.getElementById('staf-email').value.trim();
  const pw = document.getElementById('staf-password').value;
  showSpinner(true);
  const { data, error } = await sb.auth.signInWithPassword({ email, password: pw });
  showSpinner(false);
  if(error){ showErr('login-error','Email atau password salah.'); return; }
  await afterLogin();
}

/* ---------- LOGIN — NASABAH ---------- */
async function doLoginNasabah(){
  hideErr('login-error');
  const idVal = document.getElementById('nas-id').value.trim();
  const pw = document.getElementById('nas-password').value;
  showSpinner(true);
  const { data: email, error: lookupErr } = await sb.rpc('get_login_email', { p_identifier: idVal });
  if(lookupErr || !email){ showSpinner(false); showErr('login-error','No. rekening / email tidak ditemukan atau akun belum aktif.'); return; }
  const { error } = await sb.auth.signInWithPassword({ email, password: pw });
  showSpinner(false);
  if(error){ showErr('login-error','Password salah.'); return; }
  await afterLogin();
}

async function afterLogin(){
  const { data:{session} } = await sb.auth.getSession();
  if(!session){ showErr('login-error','Gagal login.'); return; }
  const { data: profile, error } = await sb.from('profiles').select('*').eq('id', session.user.id).single();
  if(error || !profile){ showErr('login-error','Profil tidak ditemukan. Hubungi Admin.'); await sb.auth.signOut(); return; }
  if(profile.status==='pending'){ showErr('login-error','Akun Anda masih menunggu verifikasi Admin.'); await sb.auth.signOut(); return; }
  if(profile.status==='rejected'){ showErr('login-error','Pendaftaran akun Anda ditolak. Hubungi Admin.'); await sb.auth.signOut(); return; }
  currentProfile = profile;
  if(profile.must_change_password){
    pendingForceUser = profile;
    document.getElementById('screen-auth').classList.add('hidden');
    document.getElementById('screen-forcepw').classList.remove('hidden');
  } else {
    enterApp();
  }
}

async function submitForcePw(){
  const p1=document.getElementById('forcepw-1').value, p2=document.getElementById('forcepw-2').value;
  if(p1.length<6){ showErr('forcepw-error','Password minimal 6 karakter.'); return; }
  if(p1!==p2){ showErr('forcepw-error','Konfirmasi password tidak cocok.'); return; }
  showSpinner(true);
  const { error: pwErr } = await sb.auth.updateUser({ password: p1 });
  if(pwErr){ showSpinner(false); showErr('forcepw-error', pwErr.message); return; }
  await sb.from('profiles').update({ must_change_password:false }).eq('id', currentProfile.id);
  currentProfile.must_change_password = false;
  showSpinner(false);
  document.getElementById('screen-forcepw').classList.add('hidden');
  enterApp();
}

/* ---------- REGISTER NASABAH (mandiri, online) ---------- */
function looksLikeEmail(s){ return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s||''); }
async function submitRegistration(){
  hideErr('register-error');
  const nama = document.getElementById('reg-nama').value.trim();
  const rt = document.getElementById('reg-rt').value.trim();
  const rw = document.getElementById('reg-rw').value.trim();
  const alamat = document.getElementById('reg-alamat').value.trim();
  const kontak = document.getElementById('reg-kontak').value.trim();
  const pw = document.getElementById('reg-password').value;
  if(!nama || !alamat){ showErr('register-error','Nama dan Alamat Rumah wajib diisi.'); return; }
  if(pw.length<6){ showErr('register-error','Password minimal 6 karakter.'); return; }

  const emailUntukAuth = looksLikeEmail(kontak) ? kontak : `reg-${Math.random().toString(36).slice(2,10)}@zetbank.local`;
  showSpinner(true);
  const { data, error } = await sb.auth.signUp({ email: emailUntukAuth, password: pw });
  if(error){ showSpinner(false); showErr('register-error', error.message); return; }
  const { error: profileErr } = await sb.from('profiles').insert({
    id: data.user.id, role:'nasabah', name:nama, rt, rw, alamat, kontak: kontak||null,
    status:'pending', must_change_password:false
  });
  await sb.auth.signOut(); // jangan biarkan tetap login sebelum diverifikasi
  showSpinner(false);
  if(profileErr){ showErr('register-error', profileErr.message); return; }
  document.getElementById('register-ok').textContent = '✓ Pendaftaran terkirim. Menunggu verifikasi Admin.';
  document.getElementById('register-ok').classList.remove('hidden');
  ['reg-nama','reg-rt','reg-rw','reg-alamat','reg-kontak','reg-password'].forEach(id=>document.getElementById(id).value='');
}

/* ---------- LOGOUT ---------- */
async function doLogout(){
  await sb.auth.signOut();
  currentProfile = null;
  document.getElementById('app').classList.remove('active');
  document.getElementById('screen-auth').classList.remove('hidden');
  document.getElementById('staf-password').value=''; document.getElementById('nas-password').value='';
}

/* ================================================================
   PERAN & MASUK APLIKASI
   ================================================================ */
function role(){ return currentProfile ? currentProfile.role : null; }
function isStaff(){ return role()==='manager'||role()==='admin'; }
function isManager(){ return role()==='manager'; }
function roleLabel(r){ return {manager:'MANAJER (SUPER ADMIN)', admin:'ADMIN', nasabah:'NASABAH'}[r] || '-'; }

async function enterApp(){
  document.getElementById('screen-auth').classList.add('hidden');
  document.getElementById('screen-forcepw').classList.add('hidden');
  document.getElementById('app').classList.add('active');
  document.getElementById('who-name').textContent = currentProfile.name;
  document.getElementById('who-role').textContent = roleLabel(currentProfile.role);
  document.getElementById('privacy-note').classList.toggle('hidden', isStaff());
  document.getElementById('trx-tanggal').value = todayStr();
  document.getElementById('trx-jam').value = nowTimeStr();
  await buildToolbar();
  await renderLedger();
}

async function buildToolbar(){
  const tb = document.getElementById('toolbar');
  const items = [];
  if(isStaff()){
    items.push(`<button class="tbtn" onclick="openModal('modal-bukutabungan')">🔍 Cek Buku Tabungan</button>`);
    const { count } = await sb.from('profiles').select('id',{count:'exact',head:true}).eq('role','nasabah').eq('status','pending');
    items.push(`<button class="tbtn" onclick="openModal('modal-verifikasi')">🧾 Verifikasi Pendaftaran ${count>0?`<span class="badge">${count}</span>`:''}</button>`);
    items.push(`<button class="tbtn" onclick="openModal('modal-nominal')">⚙ Kelola Nominal Uang</button>`);
  }
  if(isManager()) items.push(`<button class="tbtn" onclick="openModal('modal-pengguna')">👷 Kelola Pengguna</button>`);
  if(isStaff()) items.push(`<button class="tbtn primary" onclick="bukaModalTransaksiBaru()">＋ Tambah Transaksi</button>`);
  items.push(`<span class="spacer"></span>`);
  if(isStaff()){
    items.push(`<button class="tbtn" onclick="unduhBackupJSON()">{ } Unduh Backup JSON</button>`);
    items.push(`<button class="tbtn" onclick="eksporExcel()">X Ekspor Excel</button>`);
    items.push(`<button class="tbtn" onclick="eksporPDF()">P Ekspor PDF</button>`);
  }
  tb.innerHTML = items.join('');
}

function openModal(id){
  document.getElementById(id).classList.add('active');
  if(id==='modal-nominal') renderNominalTable();
  if(id==='modal-pengguna') switchUserTab('nasabah');
  if(id==='modal-verifikasi') renderVerifikasiTable();
  if(id==='modal-bukutabungan'){ btSelectedNasabah=null; document.getElementById('bt-search').value=''; document.getElementById('bt-detail').innerHTML='<p class="hint">Pilih nasabah untuk menampilkan riwayat dan saldo buku tabungannya.</p>'; document.getElementById('bt-download').disabled=true; }
}
function closeModal(id){ document.getElementById(id).classList.remove('active'); }

/* ================================================================
   SALDO BERJALAN (running balance) — dihitung dari data transaksi
   ================================================================ */
function computeSaldoMap(trxArray){
  const byNasabah = {};
  trxArray.forEach(t=>{ (byNasabah[t.nasabah_id] = byNasabah[t.nasabah_id] || []).push(t); });
  const map = {};
  Object.keys(byNasabah).forEach(nid=>{
    const list = byNasabah[nid].slice().sort((a,b)=> (a.tanggal+a.jam+(a.created_at||'')).localeCompare(b.tanggal+b.jam+(b.created_at||'')));
    let saldo=0;
    list.forEach(t=>{ saldo += t.tipe==='setor'? Number(t.nominal) : -Number(t.nominal); map[t.id]=saldo; });
  });
  return map;
}

/* ================================================================
   BUKU BESAR TRANSAKSI
   ================================================================ */
function resetFilters(){
  document.getElementById('f-from').value=''; document.getElementById('f-to').value='';
  document.getElementById('f-search').value=''; document.getElementById('f-tipe').value='';
  renderLedger();
}

async function renderLedger(){
  showSpinner(true);
  const [{data:profiles}, {data:allTrx}] = await Promise.all([
    sb.from('profiles').select('*'),
    sb.from('transactions').select('*')
  ]);
  showSpinner(false);
  const profs = profiles||[]; const trx = allTrx||[];
  const saldoMap = computeSaldoMap(trx);
  renderSaldoTotal(profs, trx);

  const from = document.getElementById('f-from').value;
  const to = document.getElementById('f-to').value;
  const search = document.getElementById('f-search').value.trim().toLowerCase();
  const tipe = document.getElementById('f-tipe').value;
  let list = trx.slice();
  if(from) list = list.filter(t=>t.tanggal>=from);
  if(to) list = list.filter(t=>t.tanggal<=to);
  if(tipe) list = list.filter(t=>t.tipe===tipe);
  if(search){
    list = list.filter(t=>{
      const nas = profs.find(p=>p.id===t.nasabah_id);
      const name = nas? nas.name.toLowerCase():''; const rek = nas? (nas.no_rekening||'').toLowerCase():'';
      return name.includes(search) || rek.includes(search);
    });
  }
  list.sort((a,b)=> (b.tanggal+b.jam).localeCompare(a.tanggal+a.jam));

  const showNominal = isStaff() || role()==='nasabah';
  const body = document.getElementById('ledger-body');
  body.innerHTML = list.map(t=>{
    const nas = profs.find(p=>p.id===t.nasabah_id);
    const nasName = nas? nas.name : '(nasabah dihapus)';
    const nasRek = nas? nas.no_rekening : '-';
    const nominalCell = showNominal ? `<span class="${t.tipe==='setor'?'amt-in':'amt-out'}">${t.tipe==='setor'?'+':'-'} ${rupiah(t.nominal)}</span>` : `<span class="amt-hidden">••••••</span>`;
    const saldoCell = showNominal ? `<span class="mono">${rupiah(saldoMap[t.id]||0)}</span>` : `<span class="amt-hidden">••••••</span>`;
    return `<tr>
      <td class="mono">${formatTgl(t.tanggal)}<br><span style="color:var(--zb-muted);font-size:11px;">${t.jam}</span></td>
      <td>${nasName}</td><td><span class="acct">${nasRek}</span></td>
      <td><span class="tag ${t.tipe==='setor'?'tag-setor':'tag-tarik'}">${t.tipe==='setor'?'Setor':'Tarik'}</span></td>
      <td>${nominalCell}</td><td>${saldoCell}</td>
      <td style="color:var(--zb-muted);">${t.keterangan||'-'}</td>
      <td style="white-space:nowrap;">${renderAksiTransaksi(t)}</td>
    </tr>`;
  }).join('');
  document.getElementById('ledger-count').textContent = list.length+' catatan';
  document.getElementById('ledger-empty').classList.toggle('hidden', list.length>0);
}

function renderSaldoTotal(profs, trx){
  const box = document.getElementById('saldo-total-box');
  if(!isStaff()){ box.classList.add('hidden'); return; }
  box.classList.remove('hidden');
  const saldoMap = computeSaldoMap(trx);
  const nasabahAktif = new Set(profs.filter(p=>p.role==='nasabah' && p.status==='active').map(p=>p.id));
  let total = 0;
  trx.forEach(t=>{ if(nasabahAktif.has(t.nasabah_id)) total = saldoMap[t.id] !== undefined ? total : total; });
  // Jumlahkan saldo akhir tiap nasabah (nilai terakhir tiap nasabah pada saldoMap)
  const lastPerNasabah = {};
  trx.slice().sort((a,b)=>(a.tanggal+a.jam+(a.created_at||'')).localeCompare(b.tanggal+b.jam+(b.created_at||''))).forEach(t=>{
    if(nasabahAktif.has(t.nasabah_id)) lastPerNasabah[t.nasabah_id] = saldoMap[t.id];
  });
  total = Object.values(lastPerNasabah).reduce((s,v)=>s+v,0);
  document.getElementById('saldo-total-val').textContent = rupiah(total);
}

function renderAksiTransaksi(t){
  if(isManager()) return `<button class="link-btn" onclick='bukaEditTransaksi(${JSON.stringify(t.id)})'>Edit</button> · <button class="link-btn" style="color:var(--zb-red);" onclick='hapusTransaksi(${JSON.stringify(t.id)}, ${t.nominal})'>Hapus</button>`;
  if(role()==='admin') return `<button class="link-btn" style="color:var(--zb-red);" onclick='hapusTransaksi(${JSON.stringify(t.id)}, ${t.nominal})'>Hapus</button>`;
  return '';
}

async function hapusTransaksi(id, nominal){
  if(Number(nominal) !== 0){
    alert('Transaksi tidak dapat dihapus karena nominal transaksi bukan Rp 0,-.\n\nUntuk menghapus transaksi ini, nominal harus diubah menjadi Rp 0 terlebih dahulu (hanya dapat dilakukan oleh Manajer melalui tombol Edit).');
    return;
  }
  if(!confirm('Hapus transaksi ini? Tindakan tidak dapat dibatalkan.')) return;
  showSpinner(true);
  const { error } = await sb.from('transactions').delete().eq('id', id);
  showSpinner(false);
  if(error){ alert('Gagal menghapus: '+error.message); return; }
  renderLedger();
}

/* ================================================================
   TAMBAH TRANSAKSI
   ================================================================ */
let trxLastSavedId = null;
let _nasabahCache = [];

async function bukaModalTransaksiBaru(){
  document.getElementById('trx-nasabah-search').value=''; document.getElementById('trx-nasabah-id').value='';
  document.getElementById('trx-tanggal').value = todayStr(); document.getElementById('trx-jam').value = nowTimeStr();
  document.getElementById('trx-nominal').value=''; document.getElementById('trx-nominal-preset').value='';
  document.getElementById('trx-keterangan').value='';
  setTrxTipe('setor'); updateTrxTotal(); hideErr('trx-error');
  document.getElementById('trx-ok').classList.add('hidden'); trxLastSavedId = null;
  showSpinner(true);
  const { data } = await sb.from('profiles').select('*').eq('role','nasabah').eq('status','active').order('name');
  _nasabahCache = data||[];
  await renderNominalSelect();
  showSpinner(false);
  openModal('modal-transaksi');
}
function renderNasabahAC(q){
  const filtered = q? _nasabahCache.filter(u=> u.name.toLowerCase().includes(q.toLowerCase()) || (u.no_rekening||'').toLowerCase().includes(q.toLowerCase())) : _nasabahCache;
  const box = document.getElementById('trx-nasabah-ac');
  box.innerHTML = filtered.slice(0,30).map(u=>`<div class="ac-item" onclick='pilihNasabahTrx(${JSON.stringify(u.id)},${JSON.stringify(u.name)},${JSON.stringify(u.no_rekening)})'>${u.name} <span class="acct">${u.no_rekening}</span></div>`).join('') || `<div class="ac-item" style="color:var(--zb-muted);">Tidak ditemukan</div>`;
  box.classList.remove('hidden');
}
function pilihNasabahTrx(id,name,rek){
  document.getElementById('trx-nasabah-id').value = id;
  document.getElementById('trx-nasabah-search').value = `${name} (${rek})`;
  document.getElementById('trx-nasabah-ac').classList.add('hidden');
}
function setTrxTipe(tipe){
  document.getElementById('trx-tipe').value = tipe;
  document.getElementById('trx-tipe-setor').style.background = tipe==='setor'? 'var(--zb-green-soft)':'';
  document.getElementById('trx-tipe-setor').style.borderColor = tipe==='setor'? 'var(--zb-green)':'';
  document.getElementById('trx-tipe-tarik').style.background = tipe==='tarik'? 'var(--zb-red-soft)':'';
  document.getElementById('trx-tipe-tarik').style.borderColor = tipe==='tarik'? 'var(--zb-red)':'';
}
async function renderNominalSelect(){
  const { data } = await sb.from('nominal_presets').select('*').order('value');
  const sel = document.getElementById('trx-nominal-preset');
  sel.innerHTML = `<option value="">Pilih nominal cepat… (opsional)</option>` + (data||[]).map(p=>`<option value="${p.value}">${p.label} — ${rupiah(p.value)}</option>`).join('');
}
function applyNominalPreset(){ const v=document.getElementById('trx-nominal-preset').value; if(v){ document.getElementById('trx-nominal').value=v; updateTrxTotal(); } }
function updateTrxTotal(){ const raw=document.getElementById('trx-nominal').value.replace(/[^\d]/g,''); document.getElementById('trx-total').textContent = rupiah(Number(raw)||0); }

async function simpanTransaksi(keepOpen){
  hideErr('trx-error');
  const nasabah_id = document.getElementById('trx-nasabah-id').value;
  const tanggal = document.getElementById('trx-tanggal').value;
  const jam = document.getElementById('trx-jam').value;
  const tipe = document.getElementById('trx-tipe').value;
  const nominal = Number(document.getElementById('trx-nominal').value.replace(/[^\d]/g,''));
  const keterangan = document.getElementById('trx-keterangan').value.trim();
  if(!nasabah_id){ showErr('trx-error','Pilih nasabah terlebih dahulu.'); return; }
  if(!tanggal || !jam){ showErr('trx-error','Tanggal dan jam wajib diisi.'); return; }
  if(!nominal || nominal<=0){ showErr('trx-error','Nominal harus lebih dari 0.'); return; }

  showSpinner(true);
  const { data, error } = await sb.from('transactions').insert({
    nasabah_id, tanggal, jam, tipe, nominal, keterangan, created_by: currentProfile.id
  }).select().single();
  showSpinner(false);
  if(error){ showErr('trx-error', error.message); return; }
  trxLastSavedId = data.id;
  document.getElementById('trx-ok').classList.remove('hidden');
  await renderLedger(); await buildToolbar();
  if(keepOpen){
    document.getElementById('trx-nominal').value=''; document.getElementById('trx-nominal-preset').value='';
    document.getElementById('trx-keterangan').value=''; document.getElementById('trx-tanggal').value=todayStr(); document.getElementById('trx-jam').value=nowTimeStr();
    updateTrxTotal();
  } else { closeModal('modal-transaksi'); }
}

async function cetakStruk(){
  if(!trxLastSavedId){ alert('Simpan transaksi terlebih dahulu (tekan OK) sebelum mencetak struk.'); return; }
  const { data: t } = await sb.from('transactions').select('*').eq('id', trxLastSavedId).single();
  const { data: nas } = await sb.from('profiles').select('*').eq('id', t.nasabah_id).single();
  const w = window.open('', '_blank', 'width=380,height=600');
  w.document.write(`<html><head><title>Struk ZET BANK</title>
    <style>body{font-family:'Courier New',monospace;font-size:13px;padding:16px;color:#111;}h2{text-align:center;margin:0 0 2px;}.c{text-align:center;}hr{border:none;border-top:1px dashed #333;margin:10px 0;}table{width:100%;}td{padding:3px 0;}</style></head><body>
    <h2>ZET BANK</h2><div class="c">Bukti Transaksi</div><hr>
    <table>
      <tr><td>Tanggal</td><td style="text-align:right;">${formatTgl(t.tanggal)} ${t.jam}</td></tr>
      <tr><td>Nasabah</td><td style="text-align:right;">${nas?nas.name:'-'}</td></tr>
      <tr><td>No. Rekening</td><td style="text-align:right;">${nas?nas.no_rekening:'-'}</td></tr>
      <tr><td>Tipe</td><td style="text-align:right;">${t.tipe==='setor'?'Setor Tunai':'Tarik Saldo'}</td></tr>
      <tr><td>Nominal</td><td style="text-align:right;">${rupiah(t.nominal)}</td></tr>
      ${t.keterangan?`<tr><td>Ket.</td><td style="text-align:right;">${t.keterangan}</td></tr>`:''}
    </table><hr><div class="c">Terima kasih<br>Simpan struk ini sebagai bukti</div>
    <script>window.print();<\/script></body></html>`);
  w.document.close();
}

/* ================================================================
   EDIT TRANSAKSI (khusus Manajer)
   ================================================================ */
async function bukaEditTransaksi(id){
  if(!isManager()) return;
  const { data: t } = await sb.from('transactions').select('*').eq('id', id).single();
  const { data: nas } = await sb.from('profiles').select('*').eq('id', t.nasabah_id).single();
  document.getElementById('edit-trx-id').value = t.id;
  document.getElementById('edit-trx-nasabah-info').innerHTML = `Nasabah: <b>${nas?nas.name:'(nasabah dihapus)'}</b> <span class="acct">${nas?nas.no_rekening:'-'}</span>`;
  document.getElementById('edit-trx-tanggal').value = t.tanggal;
  document.getElementById('edit-trx-jam').value = t.jam;
  document.getElementById('edit-trx-nominal').value = t.nominal;
  document.getElementById('edit-trx-keterangan').value = t.keterangan||'';
  setEditTrxTipe(t.tipe); hideErr('edit-trx-error');
  openModal('modal-edit-transaksi');
}
function setEditTrxTipe(tipe){
  document.getElementById('edit-trx-tipe').value = tipe;
  document.getElementById('edit-trx-tipe-setor').style.background = tipe==='setor'? 'var(--zb-green-soft)':'';
  document.getElementById('edit-trx-tipe-setor').style.borderColor = tipe==='setor'? 'var(--zb-green)':'';
  document.getElementById('edit-trx-tipe-tarik').style.background = tipe==='tarik'? 'var(--zb-red-soft)':'';
  document.getElementById('edit-trx-tipe-tarik').style.borderColor = tipe==='tarik'? 'var(--zb-red)':'';
}
async function simpanEditTransaksi(){
  if(!isManager()) return;
  hideErr('edit-trx-error');
  const id = document.getElementById('edit-trx-id').value;
  const tanggal = document.getElementById('edit-trx-tanggal').value;
  const jam = document.getElementById('edit-trx-jam').value;
  const tipe = document.getElementById('edit-trx-tipe').value;
  const nominal = Number(document.getElementById('edit-trx-nominal').value.replace(/[^\d]/g,''));
  const keterangan = document.getElementById('edit-trx-keterangan').value.trim();
  if(!tanggal || !jam){ showErr('edit-trx-error','Tanggal dan jam wajib diisi.'); return; }
  if(nominal<0){ showErr('edit-trx-error','Nominal tidak boleh negatif.'); return; }
  showSpinner(true);
  const { error } = await sb.from('transactions').update({ tanggal, jam, tipe, nominal, keterangan }).eq('id', id);
  showSpinner(false);
  if(error){ showErr('edit-trx-error', error.message); return; }
  closeModal('modal-edit-transaksi'); renderLedger();
}

/* ================================================================
   KELOLA NOMINAL UANG
   ================================================================ */
async function renderNominalTable(){
  const { data } = await sb.from('nominal_presets').select('*').order('value');
  document.getElementById('nominal-body').innerHTML = (data||[]).map(p=>`
    <tr><td>${p.label}</td><td>${rupiah(p.value)}</td><td><button class="link-btn" style="color:var(--zb-red);" onclick='hapusNominal(${JSON.stringify(p.id)})'>Hapus</button></td></tr>`).join('');
}
async function tambahNominal(){
  const label = document.getElementById('new-nominal-label').value.trim();
  const value = Number(document.getElementById('new-nominal-value').value.replace(/[^\d]/g,''));
  if(!label || !value){ alert('Label dan nominal wajib diisi.'); return; }
  await sb.from('nominal_presets').insert({label, value});
  document.getElementById('new-nominal-label').value=''; document.getElementById('new-nominal-value').value='';
  renderNominalTable();
}
async function hapusNominal(id){ await sb.from('nominal_presets').delete().eq('id', id); renderNominalTable(); }

/* ================================================================
   BUKU TABUNGAN NASABAH
   ================================================================ */
let btSelectedNasabah = null;
let _btNasabahCache = [];
async function renderBukuTabunganAC(q){
  if(_btNasabahCache.length===0){
    const { data } = await sb.from('profiles').select('*').eq('role','nasabah').eq('status','active').order('name');
    _btNasabahCache = data||[];
  }
  const filtered = q? _btNasabahCache.filter(u=> u.name.toLowerCase().includes(q.toLowerCase()) || (u.no_rekening||'').toLowerCase().includes(q.toLowerCase())) : _btNasabahCache;
  const box = document.getElementById('bt-ac');
  box.innerHTML = filtered.slice(0,30).map(u=>`<div class="ac-item" onclick='pilihNasabahBT(${JSON.stringify(u.id)})'>${u.name} <span class="acct">${u.no_rekening}</span></div>`).join('') || `<div class="ac-item" style="color:var(--zb-muted);">Tidak ditemukan</div>`;
  box.classList.remove('hidden');
}
async function pilihNasabahBT(id){
  btSelectedNasabah = _btNasabahCache.find(u=>u.id===id);
  document.getElementById('bt-search').value = btSelectedNasabah.name;
  document.getElementById('bt-ac').classList.add('hidden');
  await renderBTDetail();
  document.getElementById('bt-download').disabled = false;
}
async function renderBTDetail(){
  if(!btSelectedNasabah){ document.getElementById('bt-detail').innerHTML = `<p class="hint">Pilih nasabah untuk menampilkan riwayat dan saldo buku tabungannya.</p>`; return; }
  showSpinner(true);
  const { data: trxAll } = await sb.from('transactions').select('*').eq('nasabah_id', btSelectedNasabah.id);
  showSpinner(false);
  const saldoMap = computeSaldoMap(trxAll||[]);
  const saldo = (trxAll||[]).reduce((s,t)=> s + (t.tipe==='setor'? Number(t.nominal): -Number(t.nominal)), 0);
  const trx = (trxAll||[]).slice().sort((a,b)=>(b.tanggal+b.jam).localeCompare(a.tanggal+a.jam));
  document.getElementById('bt-detail').innerHTML = `
    <div class="saldo-card"><div class="lbl">${btSelectedNasabah.name} · <span class="acct">${btSelectedNasabah.no_rekening}</span></div><div class="val">${rupiah(saldo)}</div></div>
    <table class="mini-table"><thead><tr><th>Tanggal</th><th>Tipe</th><th>Nominal</th><th>Saldo</th><th>Keterangan</th></tr></thead>
    <tbody>${trx.map(t=>`<tr><td>${formatTgl(t.tanggal)} ${t.jam}</td><td><span class="tag ${t.tipe==='setor'?'tag-setor':'tag-tarik'}">${t.tipe==='setor'?'Setor':'Tarik'}</span></td><td class="mono">${t.tipe==='setor'?'+':'-'} ${rupiah(t.nominal)}</td><td class="mono">${rupiah(saldoMap[t.id]||0)}</td><td>${t.keterangan||'-'}</td></tr>`).join('') || `<tr><td colspan="5" style="text-align:center;color:var(--zb-muted);">Belum ada transaksi</td></tr>`}</tbody></table>`;
}
async function unduhBukuTabunganPDF(){
  if(!btSelectedNasabah) return;
  const { data: trxAll } = await sb.from('transactions').select('*').eq('nasabah_id', btSelectedNasabah.id);
  const saldoMap = computeSaldoMap(trxAll||[]);
  const saldo = (trxAll||[]).reduce((s,t)=> s + (t.tipe==='setor'? Number(t.nominal): -Number(t.nominal)), 0);
  const trx = (trxAll||[]).slice().sort((a,b)=>(a.tanggal+a.jam).localeCompare(b.tanggal+b.jam));
  const { jsPDF } = window.jspdf; const doc = new jsPDF();
  doc.setFontSize(16); doc.text('ZET BANK — Buku Tabungan', 14, 18);
  doc.setFontSize(11);
  doc.text(`Nasabah: ${btSelectedNasabah.name}`, 14, 28);
  doc.text(`No. Rekening: ${btSelectedNasabah.no_rekening}`, 14, 34);
  doc.text(`Saldo Saat Ini: ${rupiah(saldo)}`, 14, 40);
  doc.autoTable({ startY:48, head:[['Tanggal','Jam','Tipe','Nominal','Saldo','Keterangan']],
    body: trx.map(t=>[formatTgl(t.tanggal), t.jam, t.tipe==='setor'?'Setor':'Tarik', rupiah(t.nominal), rupiah(saldoMap[t.id]||0), t.keterangan||'-']) });
  doc.save(`BukuTabungan_${btSelectedNasabah.no_rekening}.pdf`);
}

/* ================================================================
   KELOLA PENGGUNA
   ================================================================ */
function switchUserTab(tab){
  document.querySelectorAll('.subtab').forEach(b=>b.classList.remove('active'));
  document.querySelector(`[data-usertab="${tab}"]`).classList.add('active');
  ['nasabah','admin'].forEach(t=> document.getElementById('usertab-'+t).classList.toggle('hidden', t!==tab));
  if(tab==='nasabah') renderNasabahTable();
  if(tab==='admin') renderAdminTable();
}
async function renderNasabahTable(){
  const { data } = await sb.from('profiles').select('*').eq('role','nasabah').eq('status','active').order('name');
  document.getElementById('nasabah-body').innerHTML = (data||[]).map(u=>`
    <tr><td>${u.name}</td><td><span class="acct">${u.no_rekening}</span></td><td>${(u.rt||u.rw)?(u.rt+'/'+u.rw):'-'}</td><td>${u.alamat||'-'}</td>
    <td style="white-space:nowrap;">
      <button class="link-btn" onclick='bukaEditNasabah(${JSON.stringify(u.id)})'>Edit</button> ·
      <button class="link-btn" onclick='resetPassword(${JSON.stringify(u.id)}, ${JSON.stringify(u.name)})'>Reset Password</button> ·
      <button class="link-btn" style="color:var(--zb-red);" onclick='hapusUser(${JSON.stringify(u.id)}, ${JSON.stringify(u.name)})'>Hapus</button>
    </td></tr>`).join('');
}
async function tambahNasabahLangsung(){
  const nama = document.getElementById('new-nasabah-nama').value.trim();
  const rt = document.getElementById('new-nasabah-rt').value.trim();
  const rw = document.getElementById('new-nasabah-rw').value.trim();
  const alamat = document.getElementById('new-nasabah-alamat').value.trim();
  if(!nama || !alamat){ alert('Nama dan Alamat Rumah wajib diisi.'); return; }
  showSpinner(true);
  const result = await edgeCall('create_nasabah', { name:nama, rt, rw, alamat });
  showSpinner(false);
  if(result.error){ alert('Gagal: '+result.error); return; }
  ['new-nasabah-nama','new-nasabah-rt','new-nasabah-rw','new-nasabah-alamat'].forEach(id=>document.getElementById(id).value='');
  renderNasabahTable(); renderLedger();
  alert(`Nasabah "${nama}" berhasil ditambahkan.\n\nNo. Rekening: ${result.noRekening}\nPassword awal: ${result.passwordAwal} (wajib diganti saat login pertama).`);
}
async function bukaEditNasabah(id){
  if(!isManager()) return;
  const { data: u } = await sb.from('profiles').select('*').eq('id', id).single();
  document.getElementById('edit-nasabah-id').value = u.id;
  document.getElementById('edit-nasabah-norek').value = u.no_rekening;
  document.getElementById('edit-nasabah-nama').value = u.name;
  document.getElementById('edit-nasabah-rt').value = u.rt||'';
  document.getElementById('edit-nasabah-rw').value = u.rw||'';
  document.getElementById('edit-nasabah-alamat').value = u.alamat||'';
  document.getElementById('edit-nasabah-kontak').value = u.kontak||'';
  hideErr('edit-nasabah-error');
  openModal('modal-edit-nasabah');
}
async function simpanEditNasabah(){
  if(!isManager()) return;
  hideErr('edit-nasabah-error');
  const id = document.getElementById('edit-nasabah-id').value;
  const nama = document.getElementById('edit-nasabah-nama').value.trim();
  const alamat = document.getElementById('edit-nasabah-alamat').value.trim();
  if(!nama || !alamat){ showErr('edit-nasabah-error','Nama dan Alamat Rumah wajib diisi.'); return; }
  showSpinner(true);
  const { error } = await sb.from('profiles').update({
    name:nama, rt:document.getElementById('edit-nasabah-rt').value.trim(), rw:document.getElementById('edit-nasabah-rw').value.trim(),
    alamat, kontak:document.getElementById('edit-nasabah-kontak').value.trim()
  }).eq('id', id);
  showSpinner(false);
  if(error){ showErr('edit-nasabah-error', error.message); return; }
  closeModal('modal-edit-nasabah'); renderNasabahTable(); renderLedger();
}
async function resetPassword(id, name){
  if(!isManager()) return;
  if(!confirm(`Reset password "${name}" ke password awal? Wajib membuat password baru saat login berikutnya.`)) return;
  showSpinner(true);
  const result = await edgeCall('reset_password', { userId:id });
  showSpinner(false);
  if(result.error){ alert('Gagal: '+result.error); return; }
  alert(`Password "${name}" telah direset ke: ${result.passwordBaru} (wajib diganti saat login berikutnya).`);
}
async function renderAdminTable(){
  const { data } = await sb.from('profiles').select('*').in('role',['manager','admin']).order('role');
  document.getElementById('admin-body').innerHTML = (data||[]).map(u=>`
    <tr><td>${u.name}</td><td>${u.email}</td><td>${roleLabel(u.role)}</td>
    <td>${u.role==='admin'?`<button class="link-btn" onclick='resetPassword(${JSON.stringify(u.id)}, ${JSON.stringify(u.name)})'>Reset Password</button> · <button class="link-btn" style="color:var(--zb-red);" onclick='hapusUser(${JSON.stringify(u.id)}, ${JSON.stringify(u.name)})'>Hapus</button>`:''}</td></tr>`).join('');
}
async function tambahAdmin(){
  const nama = document.getElementById('new-admin-nama').value.trim();
  const email = document.getElementById('new-admin-email').value.trim().toLowerCase();
  if(!nama || !email){ alert('Nama dan email wajib diisi.'); return; }
  showSpinner(true);
  const result = await edgeCall('create_admin', { name:nama, email });
  showSpinner(false);
  if(result.error){ alert('Gagal: '+result.error); return; }
  document.getElementById('new-admin-nama').value=''; document.getElementById('new-admin-email').value='';
  renderAdminTable();
  alert(`Admin "${nama}" berhasil ditambahkan.\n\nPassword awal (tampil sekali): ${result.passwordAwal}\n\nSampaikan secara langsung/pribadi ke admin terkait.`);
}
async function hapusUser(id, name){
  if(!confirm(`Hapus "${name}"? Jika ini nasabah, seluruh riwayat transaksinya ikut terhapus permanen.`)) return;
  showSpinner(true);
  const result = await edgeCall('delete_user', { userId:id });
  showSpinner(false);
  if(result.error){ alert(result.error); return; }
  renderNasabahTable(); renderAdminTable(); buildToolbar(); renderLedger();
}

/* ================================================================
   VERIFIKASI PENDAFTARAN
   ================================================================ */
async function renderVerifikasiTable(){
  const { data } = await sb.from('profiles').select('*').eq('role','nasabah').eq('status','pending').order('created_at');
  const list = data||[];
  document.getElementById('verifikasi-body').innerHTML = list.map(u=>`
    <tr><td>${u.name}</td><td>${(u.rt||u.rw)?(u.rt+'/'+u.rw):'-'}</td><td>${u.alamat||'-'}</td><td>${u.kontak||'-'}</td>
    <td style="white-space:nowrap;">
      <button class="btn btn-gold btn-sm" onclick='verifikasiUser(${JSON.stringify(u.id)},true)'>Setujui</button>
      <button class="btn btn-ghost btn-sm" onclick='verifikasiUser(${JSON.stringify(u.id)},false)'>Tolak</button>
    </td></tr>`).join('');
  document.getElementById('verifikasi-empty').classList.toggle('hidden', list.length>0);
}
async function verifikasiUser(id, approve){
  showSpinner(true);
  const { error } = await sb.rpc('verify_registration', { p_user_id:id, p_approve:approve });
  showSpinner(false);
  if(error){ alert('Gagal: '+error.message); return; }
  renderVerifikasiTable(); buildToolbar(); renderLedger();
}

/* ================================================================
   GANTI PASSWORD (dari dalam app)
   ================================================================ */
async function submitChangePw(){
  const p1=document.getElementById('cpw-1').value, p2=document.getElementById('cpw-2').value;
  hideErr('changepw-error');
  if(p1.length<6){ showErr('changepw-error','Password minimal 6 karakter.'); return; }
  if(p1!==p2){ showErr('changepw-error','Konfirmasi password tidak cocok.'); return; }
  showSpinner(true);
  const { error } = await sb.auth.updateUser({ password: p1 });
  if(!error) await sb.from('profiles').update({ must_change_password:false }).eq('id', currentProfile.id);
  showSpinner(false);
  if(error){ showErr('changepw-error', error.message); return; }
  document.getElementById('cpw-1').value=''; document.getElementById('cpw-2').value='';
  closeModal('modal-changepw');
  alert('Password berhasil diubah.');
}

/* ================================================================
   BACKUP JSON (unduh salinan arsip — bukan alat pemulihan otomatis)
   ================================================================ */
async function unduhBackupJSON(){
  showSpinner(true);
  const [{data:profiles},{data:trx},{data:nominal}] = await Promise.all([
    sb.from('profiles').select('*'), sb.from('transactions').select('*'), sb.from('nominal_presets').select('*')
  ]);
  showSpinner(false);
  const data = { exportedAt:new Date().toISOString(), profiles, transactions:trx, nominalPresets:nominal };
  const blob = new Blob([JSON.stringify(data,null,2)], {type:'application/json'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href=url; a.download=`zetbank-backup-${todayStr()}.json`; a.click();
  URL.revokeObjectURL(url);
}

/* ================================================================
   EKSPOR EXCEL & PDF
   ================================================================ */
async function eksporExcel(){
  showSpinner(true);
  const [{data:profiles},{data:trx}] = await Promise.all([ sb.from('profiles').select('*'), sb.from('transactions').select('*') ]);
  showSpinner(false);
  const saldoMap = computeSaldoMap(trx||[]);
  const rows = (trx||[]).map(t=>{
    const nas = (profiles||[]).find(p=>p.id===t.nasabah_id);
    return { Tanggal:t.tanggal, Jam:t.jam, Nasabah:nas?nas.name:'-', 'No. Rekening':nas?nas.no_rekening:'-',
      Tipe: t.tipe==='setor'?'Setor':'Tarik', 'Nominal (Rp)':t.nominal, 'Saldo (Rp)':saldoMap[t.id]||0, Keterangan:t.keterangan||'' };
  });
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Buku Besar');
  XLSX.writeFile(wb, `zetbank-transaksi-${todayStr()}.xlsx`);
}
async function eksporPDF(){
  showSpinner(true);
  const [{data:profiles},{data:trx}] = await Promise.all([ sb.from('profiles').select('*'), sb.from('transactions').select('*') ]);
  showSpinner(false);
  const saldoMap = computeSaldoMap(trx||[]);
  const { jsPDF } = window.jspdf; const doc = new jsPDF();
  doc.setFontSize(16); doc.text('ZET BANK — Buku Besar Transaksi', 14, 18);
  doc.setFontSize(10); doc.text(`Diekspor: ${new Date().toLocaleString('id-ID')}`, 14, 24);
  doc.autoTable({ startY:30, head:[['Tanggal','Jam','Nasabah','No. Rek','Tipe','Nominal','Saldo','Keterangan']],
    body: (trx||[]).map(t=>{ const nas=(profiles||[]).find(p=>p.id===t.nasabah_id);
      return [formatTgl(t.tanggal), t.jam, nas?nas.name:'-', nas?nas.no_rekening:'-', t.tipe==='setor'?'Setor':'Tarik', rupiah(t.nominal), rupiah(saldoMap[t.id]||0), t.keterangan||'-']; }),
    styles:{fontSize:8} });
  doc.save(`zetbank-transaksi-${todayStr()}.pdf`);
}

/* ================================================================
   INIT — cek sesi login yang masih aktif saat halaman dibuka
   ================================================================ */
(async function init(){
  const { data:{session} } = await sb.auth.getSession();
  if(session){ await afterLogin(); }
})();
</script>
</body>
</html>
