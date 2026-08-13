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
<style>
/* ============================================================
   ZET BANK — DESIGN TOKENS
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

/* ---------- Ledger texture (signature motif) ---------- */
.ledger-rule{
  background-image:repeating-linear-gradient(to bottom, transparent, transparent 34px, var(--zb-line) 35px);
}
.stamp{
  display:inline-flex;align-items:center;gap:6px;
  font-family:var(--font-mono);font-size:12px;letter-spacing:.04em;
  border:1px solid var(--zb-gold);color:var(--zb-navy);
  background:var(--zb-gold-soft);
  padding:3px 9px;border-radius:999px;
  transform:rotate(-1deg);
}

/* ---------- Auth shell ---------- */
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
  background:
    linear-gradient(180deg, var(--zb-navy-2), var(--zb-navy-3));
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
.auth-tab{
  flex:1;border:none;background:transparent;padding:10px 8px;border-radius:8px;
  font-size:13.5px;font-weight:600;color:var(--zb-muted);
}
.auth-tab.active{background:var(--zb-card);color:var(--zb-navy);box-shadow:var(--zb-shadow);}
.field{margin-bottom:14px;}
.field label{display:block;font-size:12.5px;font-weight:600;color:var(--zb-muted);margin-bottom:6px;}
.field input,.field select{
  width:100%;padding:11px 12px;border:1px solid var(--zb-line);border-radius:8px;
  font-size:14.5px;background:#fff;color:var(--zb-ink);
}
.field input:focus,.field select:focus{border-color:var(--zb-gold);}
.row2{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
.hint{font-size:12px;color:var(--zb-muted);margin-top:10px;line-height:1.5;}
.btn{
  border:none;border-radius:8px;padding:11px 18px;font-weight:600;font-size:14px;
  display:inline-flex;align-items:center;justify-content:center;gap:7px;
  transition:transform .05s ease;
}
.btn:active{transform:translateY(1px);}
.btn-primary{background:var(--zb-navy);color:#fff;width:100%;}
.btn-primary:hover{background:var(--zb-navy-2);}
.btn-gold{background:var(--zb-gold);color:#fff;}
.btn-ghost{background:transparent;color:var(--zb-navy);border:1px solid var(--zb-line);}
.btn-ghost:hover{border-color:var(--zb-navy);}
.btn-danger{background:var(--zb-red);color:#fff;}
.btn-sm{padding:7px 12px;font-size:12.5px;border-radius:7px;}
.btn-block{width:100%;}
.link-btn{background:none;border:none;color:var(--zb-navy);font-weight:600;font-size:13px;text-decoration:underline;padding:0;}
.divider{display:flex;align-items:center;gap:10px;color:var(--zb-muted);font-size:12px;margin:18px 0;}
.divider::before,.divider::after{content:"";flex:1;height:1px;background:var(--zb-line);}
.error-msg{background:var(--zb-red-soft);color:var(--zb-red);font-size:12.5px;padding:9px 12px;border-radius:7px;margin-bottom:12px;}
.ok-msg{background:var(--zb-green-soft);color:var(--zb-green);font-size:12.5px;padding:9px 12px;border-radius:7px;margin-bottom:12px;}

/* ---------- App shell ---------- */
.app-shell{min-height:100vh;display:none;flex-direction:column;}
.app-shell.active{display:flex;}
.app-header{
  background:var(--zb-navy);color:#fff;padding:14px 24px;
  display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;
  position:sticky;top:0;z-index:20;
}
.app-header .brand-mark .brand-name{font-size:19px;color:#fff;}
.app-header .brand-mark .brand-badge{width:36px;height:36px;font-size:16px;}
.header-loc{font-size:11.5px;color:#9FB3CC;font-family:var(--font-mono);margin-top:1px;}
.header-left{display:flex;align-items:center;gap:18px;}
.saldo-total{
  background:rgba(184,134,59,.16);border:1px solid rgba(184,134,59,.5);
  border-radius:9px;padding:6px 14px;
}
.saldo-total .lbl{font-size:10px;text-transform:uppercase;letter-spacing:.04em;color:#E7D5B0;}
.saldo-total .val{font-family:var(--font-mono);font-size:16px;font-weight:600;color:#fff;margin-top:1px;}
.header-right{display:flex;align-items:center;gap:14px;}
.clock{font-family:var(--font-mono);font-size:13px;color:#CFE0F2;letter-spacing:.03em;}
.who{font-size:12.5px;color:#CFE0F2;text-align:right;}
.who b{color:#fff;}
.icon-btn{
  background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.16);color:#fff;
  border-radius:7px;padding:7px 10px;font-size:12px;display:inline-flex;gap:6px;align-items:center;
}
.icon-btn:hover{background:rgba(255,255,255,.16);}

.toolbar{
  background:#fff;border-bottom:1px solid var(--zb-line);
  padding:12px 24px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;
}
.toolbar .spacer{flex:1;}
.tbtn{
  background:var(--zb-paper);border:1px solid var(--zb-line);color:var(--zb-navy);
  border-radius:8px;padding:8px 12px;font-size:12.5px;font-weight:600;
  display:inline-flex;align-items:center;gap:6px;position:relative;
}
.tbtn:hover{border-color:var(--zb-navy);}
.tbtn.primary{background:var(--zb-navy);color:#fff;border-color:var(--zb-navy);}
.tbtn.primary:hover{background:var(--zb-navy-2);}
.badge{
  background:var(--zb-red);color:#fff;font-size:10px;font-weight:700;
  border-radius:999px;min-width:16px;height:16px;display:inline-flex;align-items:center;justify-content:center;
  padding:0 4px;
}

.privacy-note{
  margin:14px 24px 0;background:var(--zb-gold-soft);border:1px solid var(--zb-gold);
  color:#5A4520;font-size:12.5px;padding:9px 14px;border-radius:8px;
}

.filters{margin:14px 24px 0;background:#fff;border:1px solid var(--zb-line);border-radius:10px;padding:14px 16px;display:flex;gap:12px;flex-wrap:wrap;align-items:end;}
.filters .field{margin-bottom:0;min-width:150px;flex:1;}
.filters .field input,.filters .field select{padding:9px 10px;font-size:13px;}

.content{padding:16px 24px 40px;flex:1;}
.ledger-card{background:#fff;border:1px solid var(--zb-line);border-radius:12px;overflow:hidden;box-shadow:var(--zb-shadow);}
.ledger-head{padding:16px 18px 8px;display:flex;justify-content:space-between;align-items:baseline;}
.ledger-head h2{font-family:var(--font-display);font-size:19px;margin:0;}
.ledger-count{font-size:12px;color:var(--zb-muted);font-family:var(--font-mono);}
table{width:100%;border-collapse:collapse;font-size:13px;}
thead th{
  text-align:left;font-size:11px;text-transform:uppercase;letter-spacing:.05em;
  color:var(--zb-muted);padding:10px 18px;border-bottom:1px solid var(--zb-line);background:var(--zb-paper);
}
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

/* ---------- Modal ---------- */
.modal-backdrop{
  position:fixed;inset:0;background:rgba(11,25,48,.55);display:none;
  align-items:flex-start;justify-content:center;z-index:100;padding:40px 16px;overflow-y:auto;
}
.modal-backdrop.active{display:flex;}
.modal{
  background:#fff;border-radius:14px;width:100%;max-width:560px;box-shadow:0 40px 80px -20px rgba(0,0,0,.4);
  animation:pop .16s ease;
}
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
.pw-toggle{display:none;}
</style>
</head>
<body>

<!-- ============================================================
     AUTH SCREENS
     ============================================================ -->
<div id="screen-auth" class="auth-shell">
  <div class="passbook">
    <div class="passbook-cover">
      <div class="brand-mark">
        <div class="brand-badge">Z</div>
        <div>
          <div class="brand-name">ZET BANK</div>
          <div class="brand-tag">Buku Tabungan Digital Warga</div>
        </div>
      </div>
      <div>
        <p class="cover-quote">"Setiap setoran tercatat, setiap saldo terjaga."</p>
      </div>
      <div class="cover-meta" id="cover-clock">--:--:--</div>
    </div>

    <div class="auth-panel">

      <!-- ---------- LOGIN ---------- -->
      <div id="pane-login">
        <div class="auth-tabs">
          <button class="auth-tab" data-authtab="staf" onclick="switchAuthTab('staf')">Login Staf</button>
          <button class="auth-tab active" data-authtab="nasabah" onclick="switchAuthTab('nasabah')">Nasabah</button>
        </div>

        <div id="login-error" class="error-msg hidden"></div>

        <!-- Staf login -->
        <div id="authtab-staf" class="hidden">
          <div class="field">
            <label>Email</label>
            <input type="email" id="staf-email" placeholder="nama@zetbank.id">
          </div>
          <div class="field">
            <label>Password</label>
            <input type="password" id="staf-password" placeholder="••••••••">
          </div>
          <button class="btn btn-primary" onclick="doLoginStaf()">Masuk</button>
        </div>

        <!-- Nasabah login -->
        <div id="authtab-nasabah">
          <div class="field">
            <label>No. Rekening / Email</label>
            <input type="text" id="nas-id" placeholder="No. rekening atau email">
          </div>
          <div class="field">
            <label>Password</label>
            <input type="password" id="nas-password" placeholder="••••••••">
          </div>
          <button class="btn btn-primary" onclick="doLoginNasabah()">Masuk</button>
          <div class="divider">atau</div>
          <button class="btn btn-ghost btn-block" onclick="showRegister()">Daftar Akun Baru</button>
        </div>
      </div>

      <!-- ---------- REGISTER ---------- -->
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

<!-- ============================================================
     FORCE CHANGE PASSWORD (first login)
     ============================================================ -->
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
      <p class="hint">Minimal 6 karakter. Password lama tidak bisa dipakai lagi setelah ini.</p>
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
        <div>
          <div class="brand-name">ZET BANK</div>
          <div class="header-loc" id="header-loc">Bank Keuangan Warga</div>
        </div>
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
    <div class="field">
      <label>Tipe</label>
      <select id="f-tipe" onchange="renderLedger()">
        <option value="">Semua Tipe</option>
        <option value="setor">Setor</option>
        <option value="tarik">Tarik</option>
      </select>
    </div>
    <button class="btn btn-ghost btn-sm" onclick="resetFilters()">Reset Filter</button>
  </div>

  <div class="content">
    <div class="ledger-card">
      <div class="ledger-head">
        <h2>Buku Besar Transaksi</h2>
        <span class="ledger-count" id="ledger-count">0 catatan</span>
      </div>
      <div style="overflow-x:auto;">
        <table>
          <thead>
            <tr>
              <th>Tanggal / Jam</th>
              <th>Nasabah</th>
              <th>No. Rek</th>
              <th>Tipe</th>
              <th>Nominal</th>
              <th>Saldo</th>
              <th>Keterangan</th>
              <th></th>
            </tr>
          </thead>
          <tbody id="ledger-body"></tbody>
        </table>
      </div>
      <div id="ledger-empty" class="empty-state hidden">
        <div class="big">Belum ada transaksi</div>
        <div>Transaksi yang tercatat akan muncul di sini.</div>
      </div>
    </div>
    <div class="footer-credit">ZET BANK — Manajemen Keuangan Warga · Prototipe demo (data tersimpan lokal di browser ini)</div>
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
        <input type="text" id="trx-nasabah-search" placeholder="Ketik untuk mencari, atau klik untuk melihat semua nasabah." onclick="renderNasabahAC('')" oninput="renderNasabahAC(this.value)">
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
          <button type="button" class="btn btn-ghost" id="trx-tipe-setor" style="flex:1;" onclick="setTrxTipe('setor')">Setor Tunai <span style="font-weight:400;font-size:11px;">(nasabah menabung)</span></button>
          <button type="button" class="btn btn-ghost" id="trx-tipe-tarik" style="flex:1;" onclick="setTrxTipe('tarik')">Tarik Saldo <span style="font-weight:400;font-size:11px;">(nasabah mencairkan)</span></button>
        </div>
        <input type="hidden" id="trx-tipe" value="setor">
      </div>

      <div class="field">
        <label>Nominal (Rp)</label>
        <select id="trx-nominal-preset" onchange="applyNominalPreset()" style="margin-bottom:8px;">
          <option value="">Pilih nominal cepat… (opsional)</option>
        </select>
        <input type="text" id="trx-nominal" placeholder="Masukkan nominal, mis. 50000" oninput="updateTrxTotal()">
      </div>

      <div class="total-box">
        <span class="lbl">Total Nominal</span>
        <span class="val" id="trx-total">Rp 0</span>
      </div>

      <div class="field" style="margin-top:14px;">
        <label>Keterangan (opsional)</label>
        <input type="text" id="trx-keterangan" placeholder="Catatan tambahan">
      </div>

      <p class="hint">Tekan OK untuk merekam transaksi ini &amp; langsung tambah transaksi baru untuk nasabah yang sama.</p>
    </div>
    <div class="modal-foot">
      <button class="btn btn-ghost" onclick="closeModal('modal-transaksi')">Batal</button>
      <button class="btn btn-ghost" id="btn-cetak-struk" onclick="cetakStruk()">🖨️ Cetak Struk</button>
      <button class="btn btn-gold" onclick="simpanTransaksi(true)">OK</button>
      <button class="btn btn-primary" onclick="simpanTransaksi(false)">Selesai &amp; Tutup</button>
    </div>
  </div>
</div>

<!-- ============================================================
     MODAL: EDIT TRANSAKSI (khusus Manajer / Super Admin)
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

      <div class="field">
        <label>Nominal (Rp)</label>
        <input type="text" id="edit-trx-nominal" placeholder="Masukkan nominal, mis. 50000">
      </div>

      <div class="field">
        <label>Keterangan (opsional)</label>
        <input type="text" id="edit-trx-keterangan" placeholder="Catatan tambahan">
      </div>

      <p class="hint">Perubahan ini hanya dapat dilakukan oleh Manajer (Super Admin) dan akan langsung memperbarui buku besar transaksi.</p>
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
      <p class="hint" style="margin-top:0;">Daftar nominal cepat yang muncul sebagai pilihan saat mencatat transaksi. Hanya Manajer dan Admin yang dapat mengubah daftar ini.</p>
      <table class="mini-table">
        <thead><tr><th>Label</th><th>Nominal (Rp)</th><th></th></tr></thead>
        <tbody id="nominal-body"></tbody>
      </table>
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
      <div id="bt-detail">
        <p class="hint">Pilih nasabah untuk menampilkan riwayat dan saldo buku tabungannya.</p>
      </div>
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
        <button class="subtab" data-usertab="backup" onclick="switchUserTab('backup')">Backup Otomatis</button>
      </div>

      <!-- Nasabah -->
      <div id="usertab-nasabah">
        <p class="hint" style="margin-top:0;">Input offline oleh Super Admin berdasarkan surat kuasa nasabah — langsung aktif tanpa antre verifikasi.</p>
        <table class="mini-table">
          <thead><tr><th>Nama</th><th>No. Rek</th><th>RT/RW</th><th>Alamat</th><th></th></tr></thead>
          <tbody id="nasabah-body"></tbody>
        </table>
        <div class="row2" style="margin-top:16px;">
          <div class="field"><label>Nama Nasabah Baru</label><input type="text" id="new-nasabah-nama"></div>
        </div>
        <div class="row2">
          <div class="field"><label>RT (opsional)</label><input type="text" id="new-nasabah-rt"></div>
          <div class="field"><label>RW (opsional)</label><input type="text" id="new-nasabah-rw"></div>
        </div>
        <div class="field"><label>Alamat Rumah *wajib</label><input type="text" id="new-nasabah-alamat"></div>
        <button class="btn btn-gold btn-sm" onclick="tambahNasabahLangsung()">Tambah</button>
        <p class="hint">Password awal dibuat otomatis oleh sistem: <b>123456</b>. Password ini hanya berlaku sekali — nasabah wajib menggantinya saat login pertama.</p>
      </div>

      <!-- Admin -->
      <div id="usertab-admin" class="hidden">
        <table class="mini-table">
          <thead><tr><th>Nama</th><th>Email</th><th>Peran</th><th></th></tr></thead>
          <tbody id="admin-body"></tbody>
        </table>
        <div class="row2" style="margin-top:16px;">
          <div class="field"><label>Nama Admin Baru</label><input type="text" id="new-admin-nama"></div>
          <div class="field"><label>Email</label><input type="email" id="new-admin-email"></div>
        </div>
        <button class="btn btn-gold btn-sm" onclick="tambahAdmin()">Tambah</button>
        <p class="hint">Password awal dibuat otomatis oleh sistem (acak) dan ditampilkan sekali untuk disampaikan ke admin terkait. Tombol "Kirim Reset Password" akan mengirim tautan reset ke email admin — admin membuat password barunya sendiri lewat email tersebut (simulasi pada demo ini).</p>
      </div>

      <!-- Backup -->
      <div id="usertab-backup" class="hidden">
        <p class="hint" style="margin-top:0;" id="backup-folder-status">Belum ada folder dipilih — backup otomatis akan diunduh ke folder Download bawaan browser tiap tengah malam.</p>
        <div style="display:flex;gap:10px;flex-wrap:wrap;">
          <button class="btn btn-ghost btn-sm" onclick="pilihFolderBackup()">📁 Pilih Folder di Komputer Ini</button>
          <button class="btn btn-ghost btn-sm" onclick="cobaBackupSekarang()">🧪 Coba Backup Sekarang</button>
        </div>
        <p class="hint">Backup otomatis berjalan tiap pukul 00:00 selama halaman ZET BANK ini tetap terbuka di browser (mis. dibiarkan menyala semalaman di komputer kantor). Karena alasan keamanan browser, folder perlu dipilih ulang setiap kali halaman ini dimuat ulang / dibuka baru.</p>
      </div>

    </div>
    <div class="modal-foot"><button class="btn btn-primary" onclick="closeModal('modal-pengguna')">Tutup</button></div>
  </div>
</div>

<!-- ============================================================
     MODAL: EDIT NASABAH (khusus Manajer / Super Admin)
     ============================================================ -->
<div class="modal-backdrop" id="modal-edit-nasabah">
  <div class="modal">
    <div class="modal-head"><h3>✎ Edit Data Nasabah</h3><button class="modal-close" onclick="closeModal('modal-edit-nasabah')">✕</button></div>
    <div class="modal-body">
      <div id="edit-nasabah-error" class="error-msg hidden"></div>
      <input type="hidden" id="edit-nasabah-id">
      <div class="field">
        <label>No. Rekening (tidak dapat diubah)</label>
        <input type="text" id="edit-nasabah-norek" disabled style="background:var(--zb-paper);color:var(--zb-muted);">
      </div>
      <div class="field"><label>Nama</label><input type="text" id="edit-nasabah-nama"></div>
      <div class="row2">
        <div class="field"><label>RT (opsional)</label><input type="text" id="edit-nasabah-rt"></div>
        <div class="field"><label>RW (opsional)</label><input type="text" id="edit-nasabah-rw"></div>
      </div>
      <div class="field"><label>Alamat Rumah *wajib</label><input type="text" id="edit-nasabah-alamat"></div>
      <div class="field"><label>No. HP / Email (opsional)</label><input type="text" id="edit-nasabah-kontak"></div>
      <p class="hint">No. Rekening dan Password tidak dapat diubah melalui form ini. Untuk mereset password, gunakan tombol "Reset Password" pada daftar nasabah.</p>
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
      <p class="hint" style="margin-top:0;">Daftar pendaftaran akun nasabah baru yang masuk secara online. Verifikasi untuk mengaktifkan akun (nomor rekening dibuat otomatis), atau tolak jika data tidak valid.</p>
      <table class="mini-table">
        <thead><tr><th>Nama</th><th>RT/RW</th><th>Alamat</th><th>Kontak</th><th></th></tr></thead>
        <tbody id="verifikasi-body"></tbody>
      </table>
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
      <p class="hint">Minimal 6 karakter. Anda akan tetap login setelah ini, tidak perlu masuk ulang.</p>
    </div>
    <div class="modal-foot">
      <button class="btn btn-ghost" onclick="closeModal('modal-changepw')">Batal</button>
      <button class="btn btn-primary" onclick="submitChangePw()">Simpan</button>
    </div>
  </div>
</div>

<!-- Hidden file input for JSON restore -->
<input type="file" id="file-restore" accept="application/json" class="hidden" onchange="handleRestoreFile(event)">

<script>
/* ================================================================
   ZET BANK — APP LOGIC
   ----------------------------------------------------------------
   SKEMA LOCALSTORAGE
   ----------------------------------------------------------------
   zb_users          : Array<User>
      User = {
        id, role: 'manager' | 'admin' | 'nasabah',
        name, email (staf) / noRekening (nasabah),
        passwordHash, mustChangePassword: bool,
        rt, rw, alamat, kontak,             // nasabah only
        status: 'active' | 'pending' | 'rejected'
      }

   zb_transactions   : Array<Transaction>
      Transaction = {
        id, nasabahId, tanggal, jam,
        tipe: 'setor' | 'tarik', nominal (number),
        keterangan, createdBy, createdAt (ISO)
      }

   zb_nominal_presets: Array<{ id, label, value }>

   zb_settings       : { namaBank, lokasi, backupFolderChosen: bool }

   zb_session        : { userId } | null

   NOTE KEAMANAN: password di-hash dengan fungsi hash sederhana
   (bukan bcrypt) karena ini berjalan 100% di browser tanpa server.
   Untuk penggunaan produksi sungguhan, autentikasi & penyimpanan
   sebaiknya dipindah ke backend + database yang aman.
   ================================================================ */

const LS_KEYS = {
  users:'zb_users', trx:'zb_transactions', nominal:'zb_nominal_presets',
  settings:'zb_settings', session:'zb_session'
};

function load(key, fallback){ try{ const v = localStorage.getItem(key); return v? JSON.parse(v): fallback; }catch(e){ return fallback; } }
function save(key, val){ localStorage.setItem(key, JSON.stringify(val)); }
function uid(prefix){ return prefix+'_'+Math.random().toString(36).slice(2,9)+Date.now().toString(36).slice(-4); }
function simpleHash(str){ let h=0; for(let i=0;i<str.length;i++){ h=(h<<5)-h+str.charCodeAt(i); h|=0; } return 'h'+h.toString(36); }
function rupiah(n){ return 'Rp '+ (Number(n)||0).toLocaleString('id-ID'); }
function todayStr(){ return new Date().toISOString().slice(0,10); }
function nowTimeStr(){ return new Date().toTimeString().slice(0,5); }
function genNoRekening(){
  // Nomor rekening dijamin unik & tidak pernah dipakai ulang: memakai penghitung permanen
  // di zb_settings (naik terus, tidak pernah mundur meski ada nasabah yang dihapus),
  // ditambah pengecekan tabrakan terhadap data yang ada sebagai pengaman ganda.
  const settings = load(LS_KEYS.settings, {});
  let seq = settings.lastRekeningSeq || 1000;
  const users = load(LS_KEYS.users, []);
  let candidate;
  do{
    seq++;
    candidate = 'ZB'+String(seq);
  }while(users.some(u=>u.noRekening===candidate));
  settings.lastRekeningSeq = seq;
  save(LS_KEYS.settings, settings);
  return candidate;
}

/* ---------- Seed data on first run ---------- */
function seedIfEmpty(){
  let users = load(LS_KEYS.users, null);
  if(!users){
    users = [{
      id: uid('u'), role:'manager', name:'Manajer Utama',
      email:'manajer@zetbank.id', passwordHash: simpleHash('zetbank123'),
      mustChangePassword:true, status:'active'
    }];
    save(LS_KEYS.users, users);
  }
  let nominal = load(LS_KEYS.nominal, null);
  if(!nominal){
    nominal = [
      {id:uid('n'), label:'Setoran Rutin', value:50000},
      {id:uid('n'), label:'Setoran Kecil', value:20000},
      {id:uid('n'), label:'Setoran Besar', value:100000}
    ];
    save(LS_KEYS.nominal, nominal);
  }
  let settings = load(LS_KEYS.settings, null);
  if(!settings){
    settings = {namaBank:'ZET BANK', lokasi:'Bank Keuangan Warga', backupFolderChosen:false, lastRekeningSeq:1000};
    save(LS_KEYS.settings, settings);
  }
  if(load(LS_KEYS.trx, null) === null) save(LS_KEYS.trx, []);
}
seedIfEmpty();

/* ---------- Session ---------- */
let session = load(LS_KEYS.session, null);
let pendingForceUser = null; // user object mid force-password-change

function currentUser(){
  if(!session) return null;
  return load(LS_KEYS.users, []).find(u=>u.id===session.userId) || null;
}

/* ================================================================
   CLOCK
   ================================================================ */
function tickClock(){
  const t = new Date().toLocaleTimeString('id-ID');
  const c1 = document.getElementById('cover-clock'); if(c1) c1.textContent = t;
  const c2 = document.getElementById('app-clock'); if(c2) c2.textContent = t;
}
setInterval(tickClock, 1000); tickClock();

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
function showRegister(){
  document.getElementById('pane-login').classList.add('hidden');
  document.getElementById('pane-register').classList.remove('hidden');
}
function backToLogin(){
  document.getElementById('pane-register').classList.add('hidden');
  document.getElementById('pane-login').classList.remove('hidden');
}
function hideErr(id){ document.getElementById(id).classList.add('hidden'); }
function showErr(id, msg){ const el=document.getElementById(id); el.textContent = msg; el.classList.remove('hidden'); }

/* ================================================================
   LOGIN — STAF (manager/admin)
   ================================================================ */
function doLoginStaf(){
  const email = document.getElementById('staf-email').value.trim().toLowerCase();
  const pw = document.getElementById('staf-password').value;
  const users = load(LS_KEYS.users, []);
  const u = users.find(x=> (x.role==='manager'||x.role==='admin') && x.email.toLowerCase()===email);
  if(!u || u.passwordHash !== simpleHash(pw)){
    showErr('login-error','Email atau password salah.'); return;
  }
  if(u.status !== 'active'){ showErr('login-error','Akun tidak aktif.'); return; }
  loginAs(u);
}

/* ---------- LOGIN — NASABAH ---------- */
function doLoginNasabah(){
  const idVal = document.getElementById('nas-id').value.trim().toLowerCase();
  const pw = document.getElementById('nas-password').value;
  const users = load(LS_KEYS.users, []);
  const u = users.find(x=> x.role==='nasabah' &&
    ((x.noRekening||'').toLowerCase()===idVal || (x.kontak||'').toLowerCase()===idVal));
  if(!u || u.passwordHash !== simpleHash(pw)){ showErr('login-error','No. rekening / password salah.'); return; }
  if(u.status==='pending'){ showErr('login-error','Akun Anda masih menunggu verifikasi Admin.'); return; }
  if(u.status==='rejected'){ showErr('login-error','Pendaftaran akun Anda ditolak. Hubungi Admin.'); return; }
  loginAs(u);
}

function loginAs(u){
  session = {userId:u.id}; save(LS_KEYS.session, session);
  hideErr('login-error');
  if(u.mustChangePassword){
    pendingForceUser = u;
    document.getElementById('screen-auth').classList.add('hidden');
    document.getElementById('screen-forcepw').classList.remove('hidden');
  } else {
    enterApp();
  }
}

function submitForcePw(){
  const p1 = document.getElementById('forcepw-1').value;
  const p2 = document.getElementById('forcepw-2').value;
  if(p1.length<6){ showErr('forcepw-error','Password minimal 6 karakter.'); return; }
  if(p1!==p2){ showErr('forcepw-error','Konfirmasi password tidak cocok.'); return; }
  const users = load(LS_KEYS.users, []);
  const idx = users.findIndex(x=>x.id===pendingForceUser.id);
  users[idx].passwordHash = simpleHash(p1);
  users[idx].mustChangePassword = false;
  save(LS_KEYS.users, users);
  document.getElementById('screen-forcepw').classList.add('hidden');
  enterApp();
}

/* ---------- REGISTER NASABAH ---------- */
function submitRegistration(){
  const nama = document.getElementById('reg-nama').value.trim();
  const rt = document.getElementById('reg-rt').value.trim();
  const rw = document.getElementById('reg-rw').value.trim();
  const alamat = document.getElementById('reg-alamat').value.trim();
  const kontak = document.getElementById('reg-kontak').value.trim();
  const pw = document.getElementById('reg-password').value;
  hideErr('register-error');
  if(!nama || !alamat){ showErr('register-error','Nama dan Alamat Rumah wajib diisi.'); return; }
  if(pw.length<6){ showErr('register-error','Password minimal 6 karakter.'); return; }
  const users = load(LS_KEYS.users, []);
  users.push({
    id:uid('u'), role:'nasabah', name:nama, rt, rw, alamat, kontak,
    noRekening:null, passwordHash:simpleHash(pw), mustChangePassword:false,
    status:'pending'
  });
  save(LS_KEYS.users, users);
  document.getElementById('register-ok').textContent = '✓ Pendaftaran terkirim. Menunggu verifikasi Admin.';
  document.getElementById('register-ok').classList.remove('hidden');
  ['reg-nama','reg-rt','reg-rw','reg-alamat','reg-kontak','reg-password'].forEach(id=>document.getElementById(id).value='');
}

/* ---------- LOGOUT ---------- */
function doLogout(){
  session = null; save(LS_KEYS.session, null);
  document.getElementById('app').classList.remove('active');
  document.getElementById('screen-auth').classList.remove('hidden');
  document.getElementById('staf-password').value='';
  document.getElementById('nas-password').value='';
}

/* ================================================================
   ENTER APP / ROLE-BASED UI
   ================================================================ */
function role(){
  const u = currentUser();
  return u? u.role : null;
}
function isStaff(){ return role()==='manager'||role()==='admin'; }
function isManager(){ return role()==='manager'; }

function enterApp(){
  document.getElementById('screen-auth').classList.add('hidden');
  document.getElementById('screen-forcepw').classList.add('hidden');
  document.getElementById('app').classList.add('active');
  const u = currentUser();
  document.getElementById('who-name').textContent = u? u.name : 'Pengunjung';
  document.getElementById('who-role').textContent = u? roleLabel(u.role) : 'TAMU';
  document.getElementById('privacy-note').classList.toggle('hidden', isStaff());
  renderSaldoTotal();
  buildToolbar();
  renderNominalSelect();
  document.getElementById('trx-tanggal').value = todayStr();
  document.getElementById('trx-jam').value = nowTimeStr();
  renderLedger();
  maybeScheduleBackup();
}
function roleLabel(r){ return {manager:'MANAJER (SUPER ADMIN)', admin:'ADMIN', nasabah:'NASABAH'}[r] || 'TAMU'; }

function buildToolbar(){
  const tb = document.getElementById('toolbar');
  const items = [];
  if(isStaff()){
    items.push(`<button class="tbtn" onclick="openModal('modal-bukutabungan')">🔍 Cek Buku Tabungan</button>`);
    const pendCount = load(LS_KEYS.users,[]).filter(u=>u.role==='nasabah'&&u.status==='pending').length;
    items.push(`<button class="tbtn" onclick="openModal('modal-verifikasi')">🧾 Verifikasi Pendaftaran ${pendCount>0?`<span class="badge">${pendCount}</span>`:''}</button>`);
    items.push(`<button class="tbtn" onclick="openModal('modal-nominal')">⚙ Kelola Nominal Uang</button>`);
  }
  if(isManager()){
    items.push(`<button class="tbtn" onclick="openModal('modal-pengguna')">👷 Kelola Pengguna</button>`);
  }
  if(isStaff()){
    items.push(`<button class="tbtn primary" onclick="bukaModalTransaksiBaru()">＋ Tambah Transaksi</button>`);
  }
  items.push(`<span class="spacer"></span>`);
  if(isStaff()){
    items.push(`<button class="tbtn" onclick="unduhBackupJSON()">{ } Unduh Backup JSON</button>`);
    items.push(`<button class="tbtn" onclick="document.getElementById('file-restore').click()">↺ Panggil Data JSON</button>`);
    items.push(`<button class="tbtn" onclick="eksporExcel()">X Ekspor Excel</button>`);
    items.push(`<button class="tbtn" onclick="eksporPDF()">P Ekspor PDF</button>`);
  }
  tb.innerHTML = items.join('');
}

/* ================================================================
   SALDO BERJALAN (running balance)
   ----------------------------------------------------------------
   Menghitung saldo setelah setiap transaksi, per nasabah, secara
   kronologis (tanggal+jam, lalu createdAt sebagai penentu urutan
   jika tanggal/jam sama). Dihitung ulang setiap kali dipanggil
   sehingga selalu akurat walau ada transaksi yang diedit/dihapus.
   ================================================================ */
function computeSaldoMap(){
  const trx = load(LS_KEYS.trx, []);
  const byNasabah = {};
  trx.forEach(t=>{ (byNasabah[t.nasabahId] = byNasabah[t.nasabahId] || []).push(t); });
  const map = {};
  Object.keys(byNasabah).forEach(nid=>{
    const list = byNasabah[nid].slice().sort((a,b)=> (a.tanggal+a.jam+(a.createdAt||'')).localeCompare(b.tanggal+b.jam+(b.createdAt||'')));
    let saldo = 0;
    list.forEach(t=>{ saldo += t.tipe==='setor'? t.nominal : -t.nominal; map[t.id] = saldo; });
  });
  return map;
}

/* ================================================================
   MODAL HELPERS
   ================================================================ */
function openModal(id){ document.getElementById(id).classList.add('active'); }
function closeModal(id){ document.getElementById(id).classList.remove('active'); }

/* ================================================================
   LEDGER (Buku Besar Transaksi)
   ================================================================ */
function resetFilters(){
  document.getElementById('f-from').value='';
  document.getElementById('f-to').value='';
  document.getElementById('f-search').value='';
  document.getElementById('f-tipe').value='';
  renderLedger();
}

function getFilteredTrx(){
  const from = document.getElementById('f-from').value;
  const to = document.getElementById('f-to').value;
  const search = document.getElementById('f-search').value.trim().toLowerCase();
  const tipe = document.getElementById('f-tipe').value;
  const users = load(LS_KEYS.users, []);
  let list = load(LS_KEYS.trx, []).slice();

  // Nasabah hanya melihat transaksi miliknya sendiri
  if(role()==='nasabah'){
    const u = currentUser();
    list = list.filter(t=>t.nasabahId===u.id);
  }

  if(from) list = list.filter(t=>t.tanggal>=from);
  if(to) list = list.filter(t=>t.tanggal<=to);
  if(tipe) list = list.filter(t=>t.tipe===tipe);
  if(search){
    list = list.filter(t=>{
      const nas = users.find(u=>u.id===t.nasabahId);
      const name = nas? nas.name.toLowerCase():'';
      const rek = nas? (nas.noRekening||'').toLowerCase():'';
      return name.includes(search) || rek.includes(search);
    });
  }
  return list.sort((a,b)=> (b.tanggal+b.jam).localeCompare(a.tanggal+a.jam));
}

function renderLedger(){
  const list = getFilteredTrx();
  const users = load(LS_KEYS.users, []);
  const saldoMap = computeSaldoMap();
  const body = document.getElementById('ledger-body');
  const showNominal = isStaff() || role()==='nasabah'; // nasabah hanya melihat baris miliknya sendiri (sudah difilter di getFilteredTrx)
  body.innerHTML = list.map(t=>{
    const nas = users.find(u=>u.id===t.nasabahId);
    const nasName = nas? nas.name : '(nasabah dihapus)';
    const nasRek = nas? nas.noRekening : '-';
    const nominalCell = showNominal
      ? `<span class="${t.tipe==='setor'?'amt-in':'amt-out'}">${t.tipe==='setor'?'+':'-'} ${rupiah(t.nominal)}</span>`
      : `<span class="amt-hidden">••••••</span>`;
    const saldoCell = showNominal
      ? `<span class="mono">${rupiah(saldoMap[t.id]||0)}</span>`
      : `<span class="amt-hidden">••••••</span>`;
    return `<tr>
      <td class="mono">${formatTgl(t.tanggal)}<br><span style="color:var(--zb-muted);font-size:11px;">${t.jam}</span></td>
      <td>${nasName}</td>
      <td><span class="acct">${nasRek}</span></td>
      <td><span class="tag ${t.tipe==='setor'?'tag-setor':'tag-tarik'}">${t.tipe==='setor'?'Setor':'Tarik'}</span></td>
      <td>${nominalCell}</td>
      <td>${saldoCell}</td>
      <td style="color:var(--zb-muted);">${t.keterangan||'-'}</td>
      <td style="white-space:nowrap;">${renderAksiTransaksi(t.id)}</td>
    </tr>`;
  }).join('');
  document.getElementById('ledger-count').textContent = list.length+' catatan';
  document.getElementById('ledger-empty').classList.toggle('hidden', list.length>0);
  renderSaldoTotal();
}

/* ---------- Saldo total seluruh nasabah (khusus Admin & Manajer) ---------- */
function renderSaldoTotal(){
  const box = document.getElementById('saldo-total-box');
  if(!isStaff()){ box.classList.add('hidden'); return; }
  box.classList.remove('hidden');
  const nasabahList = load(LS_KEYS.users, []).filter(u=>u.role==='nasabah' && u.status==='active');
  const total = nasabahList.reduce((sum, u)=> sum + hitungSaldo(u.id), 0);
  document.getElementById('saldo-total-val').textContent = rupiah(total);
}
function renderAksiTransaksi(id){
  if(isManager()){
    return `<button class="link-btn" onclick="bukaEditTransaksi('${id}')">Edit</button> · <button class="link-btn" style="color:var(--zb-red);" onclick="hapusTransaksi('${id}')">Hapus</button>`;
  }
  if(role()==='admin'){
    return `<button class="link-btn" style="color:var(--zb-red);" onclick="hapusTransaksi('${id}')">Hapus</button>`;
  }
  return '';
}
function formatTgl(iso){ if(!iso) return '-'; const [y,m,d]=iso.split('-'); return `${d}/${m}/${y}`; }

function hapusTransaksi(id){
  if(!(isManager() || role()==='admin')) return; // hanya Admin & Manajer yang boleh menghapus
  const list0 = load(LS_KEYS.trx, []);
  const t = list0.find(x=>x.id===id);
  if(!t) return;
  if(t.nominal !== 0){
    alert('Transaksi tidak dapat dihapus karena nominal transaksi bukan Rp 0,-.\n\nUntuk menghapus transaksi ini, nominal harus diubah menjadi Rp 0 terlebih dahulu (hanya dapat dilakukan oleh Manajer melalui tombol Edit).');
    return;
  }
  if(!confirm('Hapus transaksi ini? Tindakan tidak dapat dibatalkan.')) return;
  const list = list0.filter(t=>t.id!==id);
  save(LS_KEYS.trx, list);
  renderLedger();
}

/* ---------- Edit Transaksi (khusus Manajer / Super Admin) ---------- */
function bukaEditTransaksi(id){
  if(!isManager()) return; // hanya Manajer yang boleh mengubah data transaksi
  const t = load(LS_KEYS.trx, []).find(x=>x.id===id);
  if(!t) return;
  const nas = load(LS_KEYS.users, []).find(u=>u.id===t.nasabahId);
  document.getElementById('edit-trx-id').value = t.id;
  document.getElementById('edit-trx-nasabah-info').innerHTML = `Nasabah: <b>${nas?nas.name:'(nasabah dihapus)'}</b> <span class="acct">${nas?nas.noRekening:'-'}</span>`;
  document.getElementById('edit-trx-tanggal').value = t.tanggal;
  document.getElementById('edit-trx-jam').value = t.jam;
  document.getElementById('edit-trx-nominal').value = t.nominal;
  document.getElementById('edit-trx-keterangan').value = t.keterangan||'';
  setEditTrxTipe(t.tipe);
  hideErr('edit-trx-error');
  openModal('modal-edit-transaksi');
}
function setEditTrxTipe(tipe){
  document.getElementById('edit-trx-tipe').value = tipe;
  document.getElementById('edit-trx-tipe-setor').style.background = tipe==='setor'? 'var(--zb-green-soft)':'';
  document.getElementById('edit-trx-tipe-setor').style.borderColor = tipe==='setor'? 'var(--zb-green)':'';
  document.getElementById('edit-trx-tipe-tarik').style.background = tipe==='tarik'? 'var(--zb-red-soft)':'';
  document.getElementById('edit-trx-tipe-tarik').style.borderColor = tipe==='tarik'? 'var(--zb-red)':'';
}
function simpanEditTransaksi(){
  if(!isManager()) return;
  hideErr('edit-trx-error');
  const id = document.getElementById('edit-trx-id').value;
  const tanggal = document.getElementById('edit-trx-tanggal').value;
  const jam = document.getElementById('edit-trx-jam').value;
  const tipe = document.getElementById('edit-trx-tipe').value;
  const nominal = Number(document.getElementById('edit-trx-nominal').value.replace(/[^\d]/g,''));
  const keterangan = document.getElementById('edit-trx-keterangan').value.trim();

  if(!tanggal || !jam){ showErr('edit-trx-error','Tanggal dan jam wajib diisi.'); return; }
  if(!nominal || nominal<=0){ showErr('edit-trx-error','Nominal harus lebih dari 0.'); return; }

  const list = load(LS_KEYS.trx, []);
  const idx = list.findIndex(t=>t.id===id);
  if(idx===-1) return;
  list[idx] = {...list[idx], tanggal, jam, tipe, nominal, keterangan};
  save(LS_KEYS.trx, list);
  closeModal('modal-edit-transaksi');
  renderLedger();
}

/* ================================================================
   TAMBAH TRANSAKSI
   ================================================================ */
let trxLastSavedId = null;

function bukaModalTransaksiBaru(){
  document.getElementById('trx-nasabah-search').value='';
  document.getElementById('trx-nasabah-id').value='';
  document.getElementById('trx-tanggal').value = todayStr();
  document.getElementById('trx-jam').value = nowTimeStr();
  document.getElementById('trx-nominal').value='';
  document.getElementById('trx-nominal-preset').value='';
  document.getElementById('trx-keterangan').value='';
  setTrxTipe('setor');
  updateTrxTotal();
  hideErr('trx-error');
  document.getElementById('trx-ok').classList.add('hidden');
  trxLastSavedId = null;
  openModal('modal-transaksi');
}

function renderNasabahAC(q){
  const users = load(LS_KEYS.users, []).filter(u=>u.role==='nasabah' && u.status==='active');
  const filtered = q? users.filter(u=> u.name.toLowerCase().includes(q.toLowerCase()) || (u.noRekening||'').toLowerCase().includes(q.toLowerCase())) : users;
  const box = document.getElementById('trx-nasabah-ac');
  if(filtered.length===0){ box.innerHTML = `<div class="ac-item" style="color:var(--zb-muted);">Tidak ditemukan</div>`; box.classList.remove('hidden'); return; }
  box.innerHTML = filtered.slice(0,30).map(u=>`<div class="ac-item" onclick="pilihNasabahTrx('${u.id}','${u.name.replace(/'/g,"\\'")}','${u.noRekening}')">${u.name} <span class="acct">${u.noRekening}</span></div>`).join('');
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

function renderNominalSelect(){
  const presets = load(LS_KEYS.nominal, []);
  const sel = document.getElementById('trx-nominal-preset');
  sel.innerHTML = `<option value="">Pilih nominal cepat… (opsional)</option>` +
    presets.map(p=>`<option value="${p.value}">${p.label} — ${rupiah(p.value)}</option>`).join('');
}
function applyNominalPreset(){
  const v = document.getElementById('trx-nominal-preset').value;
  if(v){ document.getElementById('trx-nominal').value = v; updateTrxTotal(); }
}
function updateTrxTotal(){
  const raw = document.getElementById('trx-nominal').value.replace(/[^\d]/g,'');
  document.getElementById('trx-total').textContent = rupiah(Number(raw)||0);
}

function simpanTransaksi(keepOpen){
  hideErr('trx-error');
  const nasabahId = document.getElementById('trx-nasabah-id').value;
  const tanggal = document.getElementById('trx-tanggal').value;
  const jam = document.getElementById('trx-jam').value;
  const tipe = document.getElementById('trx-tipe').value;
  const nominal = Number(document.getElementById('trx-nominal').value.replace(/[^\d]/g,''));
  const keterangan = document.getElementById('trx-keterangan').value.trim();

  if(!nasabahId){ showErr('trx-error','Pilih nasabah terlebih dahulu.'); return; }
  if(!tanggal || !jam){ showErr('trx-error','Tanggal dan jam wajib diisi.'); return; }
  if(!nominal || nominal<=0){ showErr('trx-error','Nominal harus lebih dari 0.'); return; }

  const trx = {
    id: uid('t'), nasabahId, tanggal, jam, tipe, nominal, keterangan,
    createdBy: currentUser()? currentUser().id : null,
    createdAt: new Date().toISOString()
  };
  const list = load(LS_KEYS.trx, []);
  list.push(trx);
  save(LS_KEYS.trx, list);
  trxLastSavedId = trx.id;

  document.getElementById('trx-ok').classList.remove('hidden');
  renderLedger();
  buildToolbar();

  if(keepOpen){
    // simpan & langsung siap transaksi baru untuk nasabah yang sama
    document.getElementById('trx-nominal').value='';
    document.getElementById('trx-nominal-preset').value='';
    document.getElementById('trx-keterangan').value='';
    document.getElementById('trx-tanggal').value = todayStr();
    document.getElementById('trx-jam').value = nowTimeStr();
    updateTrxTotal();
  } else {
    closeModal('modal-transaksi');
  }
}

function cetakStruk(){
  if(!trxLastSavedId){ alert('Simpan transaksi terlebih dahulu (tekan OK) sebelum mencetak struk.'); return; }
  const t = load(LS_KEYS.trx, []).find(x=>x.id===trxLastSavedId);
  const nas = load(LS_KEYS.users, []).find(u=>u.id===t.nasabahId);
  const w = window.open('', '_blank', 'width=380,height=600');
  w.document.write(`
    <html><head><title>Struk ZET BANK</title>
    <style>
      body{font-family:'Courier New',monospace;font-size:13px;padding:16px;color:#111;}
      h2{text-align:center;margin:0 0 2px;} .c{text-align:center;}
      hr{border:none;border-top:1px dashed #333;margin:10px 0;}
      table{width:100%;} td{padding:3px 0;}
    </style></head><body>
      <h2>ZET BANK</h2>
      <div class="c">Bukti Transaksi</div>
      <hr>
      <table>
        <tr><td>Tanggal</td><td style="text-align:right;">${formatTgl(t.tanggal)} ${t.jam}</td></tr>
        <tr><td>Nasabah</td><td style="text-align:right;">${nas?nas.name:'-'}</td></tr>
        <tr><td>No. Rekening</td><td style="text-align:right;">${nas?nas.noRekening:'-'}</td></tr>
        <tr><td>Tipe</td><td style="text-align:right;">${t.tipe==='setor'?'Setor Tunai':'Tarik Saldo'}</td></tr>
        <tr><td>Nominal</td><td style="text-align:right;">${rupiah(t.nominal)}</td></tr>
        ${t.keterangan?`<tr><td>Ket.</td><td style="text-align:right;">${t.keterangan}</td></tr>`:''}
      </table>
      <hr>
      <div class="c">Terima kasih<br>Simpan struk ini sebagai bukti</div>
      <script>window.print();<\/script>
    </body></html>`);
  w.document.close();
}

/* ================================================================
   KELOLA NOMINAL UANG
   ================================================================ */
function renderNominalTable(){
  const presets = load(LS_KEYS.nominal, []);
  document.getElementById('nominal-body').innerHTML = presets.map(p=>`
    <tr><td>${p.label}</td><td>${rupiah(p.value)}</td>
    <td><button class="link-btn" style="color:var(--zb-red);" onclick="hapusNominal('${p.id}')">Hapus</button></td></tr>`).join('');
}
function tambahNominal(){
  const label = document.getElementById('new-nominal-label').value.trim();
  const value = Number(document.getElementById('new-nominal-value').value.replace(/[^\d]/g,''));
  if(!label || !value){ alert('Label dan nominal wajib diisi.'); return; }
  const presets = load(LS_KEYS.nominal, []);
  presets.push({id:uid('n'), label, value});
  save(LS_KEYS.nominal, presets);
  document.getElementById('new-nominal-label').value='';
  document.getElementById('new-nominal-value').value='';
  renderNominalTable(); renderNominalSelect();
}
function hapusNominal(id){
  let presets = load(LS_KEYS.nominal, []);
  presets = presets.filter(p=>p.id!==id);
  save(LS_KEYS.nominal, presets);
  renderNominalTable(); renderNominalSelect();
}

/* ================================================================
   BUKU TABUNGAN NASABAH
   ================================================================ */
let btSelectedNasabah = null;
function renderBukuTabunganAC(q){
  const users = load(LS_KEYS.users, []).filter(u=>u.role==='nasabah' && u.status==='active');
  const filtered = q? users.filter(u=> u.name.toLowerCase().includes(q.toLowerCase()) || (u.noRekening||'').toLowerCase().includes(q.toLowerCase())) : users;
  const box = document.getElementById('bt-ac');
  box.innerHTML = filtered.slice(0,30).map(u=>`<div class="ac-item" onclick="pilihNasabahBT('${u.id}')">${u.name} <span class="acct">${u.noRekening}</span></div>`).join('') || `<div class="ac-item" style="color:var(--zb-muted);">Tidak ditemukan</div>`;
  box.classList.remove('hidden');
}
function pilihNasabahBT(id){
  btSelectedNasabah = load(LS_KEYS.users, []).find(u=>u.id===id);
  document.getElementById('bt-search').value = btSelectedNasabah.name;
  document.getElementById('bt-ac').classList.add('hidden');
  renderBTDetail();
  document.getElementById('bt-download').disabled = false;
}
function hitungSaldo(nasabahId){
  const trx = load(LS_KEYS.trx, []).filter(t=>t.nasabahId===nasabahId);
  return trx.reduce((s,t)=> s + (t.tipe==='setor'? t.nominal : -t.nominal), 0);
}
function renderBTDetail(){
  if(!btSelectedNasabah){ document.getElementById('bt-detail').innerHTML = `<p class="hint">Pilih nasabah untuk menampilkan riwayat dan saldo buku tabungannya.</p>`; return; }
  const saldo = hitungSaldo(btSelectedNasabah.id);
  const saldoMap = computeSaldoMap();
  const trx = load(LS_KEYS.trx, []).filter(t=>t.nasabahId===btSelectedNasabah.id).sort((a,b)=>(b.tanggal+b.jam).localeCompare(a.tanggal+a.jam));
  document.getElementById('bt-detail').innerHTML = `
    <div class="saldo-card">
      <div class="lbl">${btSelectedNasabah.name} · <span class="acct">${btSelectedNasabah.noRekening}</span></div>
      <div class="val">${rupiah(saldo)}</div>
    </div>
    <table class="mini-table">
      <thead><tr><th>Tanggal</th><th>Tipe</th><th>Nominal</th><th>Saldo</th><th>Keterangan</th></tr></thead>
      <tbody>
        ${trx.map(t=>`<tr><td>${formatTgl(t.tanggal)} ${t.jam}</td><td><span class="tag ${t.tipe==='setor'?'tag-setor':'tag-tarik'}">${t.tipe==='setor'?'Setor':'Tarik'}</span></td><td class="mono">${t.tipe==='setor'?'+':'-'} ${rupiah(t.nominal)}</td><td class="mono">${rupiah(saldoMap[t.id]||0)}</td><td>${t.keterangan||'-'}</td></tr>`).join('') || `<tr><td colspan="5" style="text-align:center;color:var(--zb-muted);">Belum ada transaksi</td></tr>`}
      </tbody>
    </table>`;
}
function unduhBukuTabunganPDF(){
  if(!btSelectedNasabah) return;
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();
  const saldo = hitungSaldo(btSelectedNasabah.id);
  const saldoMap = computeSaldoMap();
  const trx = load(LS_KEYS.trx, []).filter(t=>t.nasabahId===btSelectedNasabah.id).sort((a,b)=>(a.tanggal+a.jam).localeCompare(b.tanggal+b.jam));
  doc.setFontSize(16); doc.text('ZET BANK — Buku Tabungan', 14, 18);
  doc.setFontSize(11);
  doc.text(`Nasabah: ${btSelectedNasabah.name}`, 14, 28);
  doc.text(`No. Rekening: ${btSelectedNasabah.noRekening}`, 14, 34);
  doc.text(`Saldo Saat Ini: ${rupiah(saldo)}`, 14, 40);
  doc.autoTable({
    startY: 48,
    head: [['Tanggal','Jam','Tipe','Nominal','Saldo','Keterangan']],
    body: trx.map(t=>[formatTgl(t.tanggal), t.jam, t.tipe==='setor'?'Setor':'Tarik', rupiah(t.nominal), rupiah(saldoMap[t.id]||0), t.keterangan||'-'])
  });
  doc.save(`BukuTabungan_${btSelectedNasabah.noRekening}.pdf`);
}

/* ================================================================
   KELOLA PENGGUNA
   ================================================================ */
function switchUserTab(tab){
  document.querySelectorAll('.subtab').forEach(b=>b.classList.remove('active'));
  document.querySelector(`[data-usertab="${tab}"]`).classList.add('active');
  ['nasabah','admin','backup'].forEach(t=> document.getElementById('usertab-'+t).classList.toggle('hidden', t!==tab));
  if(tab==='nasabah') renderNasabahTable();
  if(tab==='admin') renderAdminTable();
}
function renderNasabahTable(){
  const users = load(LS_KEYS.users, []).filter(u=>u.role==='nasabah' && u.status==='active');
  document.getElementById('nasabah-body').innerHTML = users.map(u=>`
    <tr><td>${u.name}</td><td><span class="acct">${u.noRekening}</span></td><td>${(u.rt||u.rw)?(u.rt+'/'+u.rw):'-'}</td><td>${u.alamat||'-'}</td>
    <td style="white-space:nowrap;">
      <button class="link-btn" onclick="bukaEditNasabah('${u.id}')">Edit</button> ·
      <button class="link-btn" onclick="resetPasswordNasabah('${u.id}')">Reset Password</button> ·
      <button class="link-btn" style="color:var(--zb-red);" onclick="hapusUser('${u.id}')">Hapus</button>
    </td></tr>`).join('');
}
function tambahNasabahLangsung(){
  const nama = document.getElementById('new-nasabah-nama').value.trim();
  const rt = document.getElementById('new-nasabah-rt').value.trim();
  const rw = document.getElementById('new-nasabah-rw').value.trim();
  const alamat = document.getElementById('new-nasabah-alamat').value.trim();
  if(!nama || !alamat){ alert('Nama dan Alamat Rumah wajib diisi.'); return; }
  const users = load(LS_KEYS.users, []);
  users.push({
    id:uid('u'), role:'nasabah', name:nama, rt, rw, alamat, kontak:'',
    noRekening: genNoRekening(), passwordHash: simpleHash('123456'), mustChangePassword:true,
    status:'active'
  });
  save(LS_KEYS.users, users);
  ['new-nasabah-nama','new-nasabah-rt','new-nasabah-rw','new-nasabah-alamat'].forEach(id=>document.getElementById(id).value='');
  renderNasabahTable();
  renderSaldoTotal();
  alert(`Nasabah "${nama}" berhasil ditambahkan.\n\nNo. Rekening: ${users[users.length-1].noRekening}\nPassword awal: 123456 (wajib diganti saat login pertama).`);
}

/* ---------- Edit data nasabah (kecuali No. Rekening & Password) ---------- */
function bukaEditNasabah(id){
  if(!isManager()) return;
  const u = load(LS_KEYS.users, []).find(x=>x.id===id);
  if(!u) return;
  document.getElementById('edit-nasabah-id').value = u.id;
  document.getElementById('edit-nasabah-norek').textContent = u.noRekening;
  document.getElementById('edit-nasabah-nama').value = u.name;
  document.getElementById('edit-nasabah-rt').value = u.rt||'';
  document.getElementById('edit-nasabah-rw').value = u.rw||'';
  document.getElementById('edit-nasabah-alamat').value = u.alamat||'';
  document.getElementById('edit-nasabah-kontak').value = u.kontak||'';
  hideErr('edit-nasabah-error');
  openModal('modal-edit-nasabah');
}
function simpanEditNasabah(){
  if(!isManager()) return;
  hideErr('edit-nasabah-error');
  const id = document.getElementById('edit-nasabah-id').value;
  const nama = document.getElementById('edit-nasabah-nama').value.trim();
  const alamat = document.getElementById('edit-nasabah-alamat').value.trim();
  if(!nama || !alamat){ showErr('edit-nasabah-error','Nama dan Alamat Rumah wajib diisi.'); return; }
  const users = load(LS_KEYS.users, []);
  const idx = users.findIndex(u=>u.id===id);
  if(idx===-1) return;
  users[idx].name = nama;
  users[idx].rt = document.getElementById('edit-nasabah-rt').value.trim();
  users[idx].rw = document.getElementById('edit-nasabah-rw').value.trim();
  users[idx].alamat = alamat;
  users[idx].kontak = document.getElementById('edit-nasabah-kontak').value.trim();
  save(LS_KEYS.users, users);
  closeModal('modal-edit-nasabah');
  renderNasabahTable();
  renderLedger();
}

/* ---------- Reset password nasabah oleh Manajer (berlaku seperti password awal) ---------- */
function resetPasswordNasabah(id){
  if(!isManager()) return;
  const u = load(LS_KEYS.users, []).find(x=>x.id===id);
  if(!u) return;
  if(!confirm(`Reset password nasabah "${u.name}" ke password awal? Nasabah akan wajib membuat password baru saat login berikutnya.`)) return;
  const users = load(LS_KEYS.users, []);
  const idx = users.findIndex(x=>x.id===id);
  users[idx].passwordHash = simpleHash('123456');
  users[idx].mustChangePassword = true;
  save(LS_KEYS.users, users);
  alert(`Password nasabah "${u.name}" telah direset ke: 123456 (wajib diganti saat login berikutnya).`);
}

function renderAdminTable(){
  const users = load(LS_KEYS.users, []).filter(u=>u.role==='manager'||u.role==='admin');
  document.getElementById('admin-body').innerHTML = users.map(u=>`
    <tr><td>${u.name}</td><td>${u.email}</td><td>${roleLabel(u.role)}</td>
    <td>${u.role==='admin'?`<button class="link-btn" onclick="kirimResetPassword('${u.id}')">Kirim Reset Password</button> · <button class="link-btn" style="color:var(--zb-red);" onclick="hapusUser('${u.id}')">Hapus</button>`:''}</td></tr>`).join('');
}
function tambahAdmin(){
  const nama = document.getElementById('new-admin-nama').value.trim();
  const email = document.getElementById('new-admin-email').value.trim().toLowerCase();
  if(!nama || !email){ alert('Nama dan email wajib diisi.'); return; }
  const users = load(LS_KEYS.users, []);
  if(users.some(u=>u.email && u.email.toLowerCase()===email)){ alert('Email sudah terdaftar.'); return; }
  const randomPw = Math.random().toString(36).slice(-8);
  users.push({ id:uid('u'), role:'admin', name:nama, email, passwordHash:simpleHash(randomPw), mustChangePassword:true, status:'active' });
  save(LS_KEYS.users, users);
  document.getElementById('new-admin-nama').value=''; document.getElementById('new-admin-email').value='';
  renderAdminTable();
  alert(`Admin "${nama}" berhasil ditambahkan.\n\nPassword awal (tampil sekali): ${randomPw}\n\nSampaikan password ini secara langsung/pribadi ke admin terkait.`);
}
function kirimResetPassword(id){
  const u = load(LS_KEYS.users, []).find(x=>x.id===id);
  alert(`(Simulasi) Tautan reset password telah "dikirim" ke ${u.email}.`);
}
function hapusUser(id){
  const users0 = load(LS_KEYS.users, []);
  const u = users0.find(x=>x.id===id);
  if(!u) return;
  if(u.role==='nasabah'){
    const saldo = hitungSaldo(u.id);
    if(saldo !== 0){
      alert(`Nasabah "${u.name}" tidak dapat dihapus karena saldo akhir bukan Rp 0,-.\n\nSaldo saat ini: ${rupiah(saldo)}\n\nSelesaikan penarikan/setoran hingga saldo Rp 0,- terlebih dahulu sebelum menghapus.`);
      return;
    }
  }
  if(!confirm('Hapus pengguna ini? Riwayat transaksinya akan tetap tersimpan.')) return;
  let users = load(LS_KEYS.users, []);
  users = users.filter(x=>x.id!==id);
  save(LS_KEYS.users, users);
  renderNasabahTable(); renderAdminTable(); buildToolbar(); renderSaldoTotal();
}

/* ---------- Backup otomatis ---------- */
let backupDirHandle = null;
async function pilihFolderBackup(){
  try{
    if(!window.showDirectoryPicker){ alert('Browser ini tidak mendukung pemilihan folder langsung. Backup akan diunduh ke folder Download.'); return; }
    backupDirHandle = await window.showDirectoryPicker();
    document.getElementById('backup-folder-status').textContent = `Folder backup dipilih: "${backupDirHandle.name}". Backup otomatis akan disimpan ke sini tiap tengah malam selama halaman ini terbuka.`;
  }catch(e){ /* user membatalkan pemilihan folder */ }
}
async function writeBackupFile(){
  const data = collectAllData();
  const filename = `zetbank-backup-${todayStr()}.json`;
  const blob = new Blob([JSON.stringify(data,null,2)], {type:'application/json'});
  if(backupDirHandle){
    try{
      const fileHandle = await backupDirHandle.getFileHandle(filename, {create:true});
      const writable = await fileHandle.createWritable();
      await writable.write(blob);
      await writable.close();
      return true;
    }catch(e){ /* fallback ke unduh biasa jika gagal menulis */ }
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href=url; a.download=filename; a.click();
  URL.revokeObjectURL(url);
  return true;
}
function cobaBackupSekarang(){ writeBackupFile().then(()=>alert('Backup berhasil dibuat: '+`zetbank-backup-${todayStr()}.json`)); }
function maybeScheduleBackup(){
  // Cek tiap menit apakah sudah lewat tengah malam sejak backup terakhir
  setInterval(()=>{
    const now = new Date();
    if(now.getHours()===0 && now.getMinutes()===0){ writeBackupFile(); }
  }, 60000);
}

/* ================================================================
   VERIFIKASI PENDAFTARAN
   ================================================================ */
function renderVerifikasiTable(){
  const users = load(LS_KEYS.users, []).filter(u=>u.role==='nasabah' && u.status==='pending');
  document.getElementById('verifikasi-body').innerHTML = users.map(u=>`
    <tr><td>${u.name}</td><td>${(u.rt||u.rw)?(u.rt+'/'+u.rw):'-'}</td><td>${u.alamat||'-'}</td><td>${u.kontak||'-'}</td>
    <td style="white-space:nowrap;">
      <button class="btn btn-gold btn-sm" onclick="verifikasiUser('${u.id}',true)">Setujui</button>
      <button class="btn btn-ghost btn-sm" onclick="verifikasiUser('${u.id}',false)">Tolak</button>
    </td></tr>`).join('');
  document.getElementById('verifikasi-empty').classList.toggle('hidden', users.length>0);
}
function verifikasiUser(id, approve){
  const users = load(LS_KEYS.users, []);
  const idx = users.findIndex(u=>u.id===id);
  if(approve){ users[idx].status='active'; users[idx].noRekening = genNoRekening(); }
  else { users[idx].status='rejected'; }
  save(LS_KEYS.users, users);
  renderVerifikasiTable(); buildToolbar(); renderSaldoTotal();
}

/* ================================================================
   GANTI PASSWORD (dari dalam app)
   ================================================================ */
function submitChangePw(){
  const p1=document.getElementById('cpw-1').value, p2=document.getElementById('cpw-2').value;
  hideErr('changepw-error');
  if(p1.length<6){ showErr('changepw-error','Password minimal 6 karakter.'); return; }
  if(p1!==p2){ showErr('changepw-error','Konfirmasi password tidak cocok.'); return; }
  const u = currentUser(); if(!u){ alert('Tidak berlaku untuk akun tamu.'); return; }
  const users = load(LS_KEYS.users, []);
  const idx = users.findIndex(x=>x.id===u.id);
  users[idx].passwordHash = simpleHash(p1);
  save(LS_KEYS.users, users);
  document.getElementById('cpw-1').value=''; document.getElementById('cpw-2').value='';
  closeModal('modal-changepw');
  alert('Password berhasil diubah.');
}

/* ================================================================
   BACKUP JSON MANUAL / RESTORE
   ================================================================ */
function collectAllData(){
  return {
    exportedAt: new Date().toISOString(),
    users: load(LS_KEYS.users, []),
    transactions: load(LS_KEYS.trx, []),
    nominalPresets: load(LS_KEYS.nominal, []),
    settings: load(LS_KEYS.settings, {})
  };
}
function unduhBackupJSON(){
  const data = collectAllData();
  const blob = new Blob([JSON.stringify(data,null,2)], {type:'application/json'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href=url; a.download=`zetbank-backup-${todayStr()}.json`; a.click();
  URL.revokeObjectURL(url);
}
function handleRestoreFile(evt){
  const file = evt.target.files[0]; if(!file) return;
  const reader = new FileReader();
  reader.onload = ()=>{
    try{
      const data = JSON.parse(reader.result);
      if(!confirm('Memanggil data JSON ini akan MENGGANTI seluruh data yang tersimpan saat ini. Lanjutkan?')) return;
      if(data.users) save(LS_KEYS.users, data.users);
      if(data.transactions) save(LS_KEYS.trx, data.transactions);
      if(data.nominalPresets) save(LS_KEYS.nominal, data.nominalPresets);
      if(data.settings) save(LS_KEYS.settings, data.settings);
      alert('Data berhasil dipanggil. Halaman akan dimuat ulang.');
      location.reload();
    }catch(e){ alert('File JSON tidak valid.'); }
  };
  reader.readAsText(file);
  evt.target.value='';
}

/* ================================================================
   EKSPOR EXCEL & PDF (Buku Besar Transaksi)
   ================================================================ */
function eksporExcel(){
  const list = getFilteredTrx();
  const users = load(LS_KEYS.users, []);
  const saldoMap = computeSaldoMap();
  const rows = list.map(t=>{
    const nas = users.find(u=>u.id===t.nasabahId);
    return {
      Tanggal: t.tanggal, Jam: t.jam,
      Nasabah: nas?nas.name:'-', 'No. Rekening': nas?nas.noRekening:'-',
      Tipe: t.tipe==='setor'?'Setor':'Tarik', 'Nominal (Rp)': t.nominal,
      'Saldo (Rp)': saldoMap[t.id]||0,
      Keterangan: t.keterangan||''
    };
  });
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Buku Besar');
  XLSX.writeFile(wb, `zetbank-transaksi-${todayStr()}.xlsx`);
}
function eksporPDF(){
  const list = getFilteredTrx();
  const users = load(LS_KEYS.users, []);
  const saldoMap = computeSaldoMap();
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();
  doc.setFontSize(16); doc.text('ZET BANK — Buku Besar Transaksi', 14, 18);
  doc.setFontSize(10); doc.text(`Diekspor: ${new Date().toLocaleString('id-ID')}`, 14, 24);
  doc.autoTable({
    startY: 30,
    head: [['Tanggal','Jam','Nasabah','No. Rek','Tipe','Nominal','Saldo','Keterangan']],
    body: list.map(t=>{
      const nas = users.find(u=>u.id===t.nasabahId);
      return [formatTgl(t.tanggal), t.jam, nas?nas.name:'-', nas?nas.noRekening:'-', t.tipe==='setor'?'Setor':'Tarik', rupiah(t.nominal), rupiah(saldoMap[t.id]||0), t.keterangan||'-'];
    }),
    styles:{fontSize:8}
  });
  doc.save(`zetbank-transaksi-${todayStr()}.pdf`);
}

/* ================================================================
   MODAL OPEN HOOKS — render data setiap kali modal dibuka
   ================================================================ */
document.getElementById('modal-nominal').addEventListener('transitionend', ()=>{});
const _origOpenModal = openModal;
openModal = function(id){
  _origOpenModal(id);
  if(id==='modal-nominal') renderNominalTable();
  if(id==='modal-pengguna'){ switchUserTab('nasabah'); }
  if(id==='modal-verifikasi') renderVerifikasiTable();
  if(id==='modal-bukutabungan'){ btSelectedNasabah=null; document.getElementById('bt-search').value=''; document.getElementById('bt-detail').innerHTML='<p class="hint">Pilih nasabah untuk menampilkan riwayat dan saldo buku tabungannya.</p>'; document.getElementById('bt-download').disabled=true; }
};

/* ================================================================
   INIT — restore session on page load
   ================================================================ */
(function init(){
  if(session && session.userId){
    const u = load(LS_KEYS.users,[]).find(x=>x.id===session.userId);
    if(u){
      if(u.mustChangePassword){
        pendingForceUser = u;
        document.getElementById('screen-auth').classList.add('hidden');
        document.getElementById('screen-forcepw').classList.remove('hidden');
      } else {
        enterApp();
      }
    }
  }
})();
</script>
</body>
</html>
