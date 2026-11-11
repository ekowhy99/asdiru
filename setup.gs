/**
 * ASDIRU v1.1 — setup.gs : pemasangan awal basis data, folder Drive, konfigurasi, dan menu kelola pengguna.
 *
 * MENU ASDIRU di Spreadsheet (khusus pemilik): tambah pengguna, atur password, nonaktifkan/aktifkan,
 * hapus pengguna, amankan password yang diketik manual, atur API Key Gemini, serta buka dan atur alamat
 * folder Drive. Tidak perlu membuka Panel Admin.
 *
 * UPGRADE dari versi sebelumnya: tempel code.gs, setup.gs, index.html terbaru lalu jalankan setupAwal() lagi.
 * Kolom baru 'jenis_kelamin' di Master_User ditambahkan otomatis (isi L atau P untuk salam Pak/Bu).
 * Kolom 'nip' di Master_User (v2.10) juga ditambahkan otomatis; NIP dicetak di blok tanda tangan dokumen.
 * Guru mapel (PAI, PJOK, dst.): isi kelas_ampu dengan beberapa kelas dan mapel_ampu dengan nama mapel
 * persis seperti di Master_Kurikulum (pisahkan koma).
 *
 * CARA PAKAI
 *  1. Buat Google Spreadsheet kosong, buka Extensions > Apps Script (script "bound").
 *  2. Tempel code.gs, setup.gs, dan index.html.
 *  3. Jalankan setupAwal() sekali dan beri izin akses.
 *     Kata sandi sementara akun contoh ditampilkan satu kali; segera ganti email & password.
 *  4. Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 *  5. Login sebagai Super Admin, lalu isi API Key Gemini dan unggah pedoman (.md) di Panel Admin.
 */

const SETUP_SCHEMA = {
  Master_User: ['id_user', 'nama_lengkap', 'email', 'password_hash', 'role', 'kelas_ampu', 'mapel_ampu', 'id_folder_drive', 'status_aktif', 'jenis_kelamin', 'nip'],
  Master_Siswa: ['id_siswa', 'nis', 'nama_siswa', 'kelas', 'jenis_kelamin', 'status_aktif'],
  Master_Jadwal: ['id_jadwal', 'hari', 'kelas', 'jam_ke', 'jam_mulai', 'jam_selesai', 'mapel', 'id_guru', 'ruang'],
  Master_Kurikulum: ['id_tp', 'fase', 'kelas', 'mapel', 'elemen', 'capaian_pembelajaran', 'tujuan_pembelajaran', 'semester', 'materi', 'atp'],
  Trx_Presensi: ['id_presensi', 'tanggal', 'kelas', 'id_siswa', 'nama_siswa', 'status', 'id_guru', 'timestamp'],
  Trx_Agenda: ['id_agenda', 'tanggal', 'id_guru', 'kelas', 'mapel', 'kegiatan', 'catatan', 'timestamp'],
  Arsip_Modul: ['id_modul', 'id_guru', 'kelas', 'mapel', 'tp', 'drive_url', 'timestamp'],
  Arsip_Asesmen: ['id_asesmen', 'id_guru', 'kelas', 'mapel', 'tp', 'drive_url', 'timestamp']
};

/** Kolom yang wajib berformat teks agar tanggal/jam/NIS tidak diubah Sheets menjadi Date/angka. */
const SETUP_TEXT_COLS = {
  Master_User: ['password_hash', 'nip'], Master_Siswa: ['nis'],
  Master_Jadwal: ['jam_ke', 'jam_mulai', 'jam_selesai'],
  Trx_Presensi: ['tanggal', 'timestamp'], Trx_Agenda: ['tanggal', 'timestamp'],
  Arsip_Modul: ['timestamp'], Arsip_Asesmen: ['timestamp']
};

function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('ASDIRU')
      .addItem('Setup awal', 'setupAwal')
      .addItem('Siapkan folder Drive guru', 'siapkanFolderGuru')
      .addItem('Buka folder Drive', 'menuBukaFolder')
      .addItem('Atur alamat folder Drive', 'menuAturFolderDrive')
      .addItem('Tambah akun dummy', 'tambahAkunDummy')
      .addSeparator()
      .addItem('Tambah pengguna', 'menuTambahPengguna')
      .addItem('Atur password pengguna', 'menuAturPassword')
      .addItem('Nonaktifkan / aktifkan pengguna', 'menuToggleStatus')
      .addItem('Hapus pengguna', 'menuHapusPengguna')
      .addItem('Amankan password yang diketik manual', 'amankanPasswordPolos')
      .addSeparator()
      .addItem('Atur API Key Gemini', 'menuApiKey')
      .addItem('Perbarui model Gemini', 'perbaruiModelGemini')
      .addItem('Cek instalasi', 'cekInstalasi')
      .addToUi();
  } catch (e) { /* bukan konteks spreadsheet */ }
}

function setupAwal() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { /* standalone */ }
  if (!ss) ss = SpreadsheetApp.create('ASDIRU - Basis Data');
  props_().setProperty('SPREADSHEET_ID', ss.getId());

  Object.keys(SETUP_SCHEMA).forEach(function (name) { setupSheet_(ss, name); });
  const def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);

  const info = [];
  if (ss.getSheetByName('Master_User').getLastRow() < 2) { info.push(seedContoh_()); tambahAkunDummy(); }
  siapkanFolderGuru();
  setupKonfigurasi_();

  const msg = 'Setup selesai.\n\nSpreadsheet: ' + ss.getUrl() + (info.length ? '\n\n' + info.join('\n') : '') +
    '\n\nLangkah berikutnya: ganti email akun contoh di Master_User, isi Master_Siswa/Jadwal/Kurikulum, lalu deploy sebagai Web App.';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('ASDIRU', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { /* lihat Logger */ }
}

function setupSheet_(ss, name) {
  const headers = SETUP_SCHEMA[name];
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getMaxColumns() < headers.length) sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
  const cur = sh.getRange(1, 1, 1, sh.getMaxColumns()).getValues()[0].map(function (v) { return String(v).trim(); });
  let n = 0;
  cur.forEach(function (v, i) { if (v !== '') n = i + 1; });
  if (n === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    headers.forEach(function (h) {                       // migrasi: kolom baru ditambahkan di kanan
      if (cur.indexOf(h) < 0) {
        n++;
        if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), 1);
        sh.getRange(1, n).setValue(h);
        cur[n - 1] = h;
      }
    });
  }
  const H = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (v) { return String(v).trim(); });
  sh.getRange(1, 1, 1, H.length).setFontWeight('bold').setBackground('#0A5BD8').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  (SETUP_TEXT_COLS[name] || []).forEach(function (c) {
    const i = H.indexOf(c) + 1;
    if (i > 0) sh.getRange(1, i, sh.getMaxRows(), 1).setNumberFormat('@');
  });
  const rule = function (list) { return SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(false).build(); };
  const max = Math.max(sh.getMaxRows() - 1, 1);
  const col = function (c) { return H.indexOf(c) + 1; };
  if (name === 'Master_User') {
    sh.getRange(2, col('role'), max, 1).setDataValidation(rule([ROLE.SUPER, ROLE.GURU, ROLE.KEPSEK]));
    sh.getRange(2, col('status_aktif'), max, 1).setDataValidation(rule(['aktif', 'nonaktif']));
    sh.getRange(2, col('jenis_kelamin'), max, 1).setDataValidation(rule(['L', 'P']));
  }
  if (name === 'Master_Siswa') sh.getRange(2, col('status_aktif'), max, 1).setDataValidation(rule(['aktif', 'nonaktif']));
  if (name === 'Master_Jadwal') sh.getRange(2, col('hari'), max, 1).setDataValidation(rule(['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu']));
  if (name === 'Trx_Presensi') sh.getRange(2, col('status'), max, 1).setDataValidation(rule(STATUS_PRESENSI));
  sh.autoResizeColumns(1, H.length);
}

function randomPass_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < 10; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

/** Data contoh agar aplikasi bisa dicoba. Ganti seluruhnya dengan data sekolah. */
function seedContoh_() {
  // [id, nama, email, role, kelas_ampu, mapel_ampu, jenis_kelamin]
  const accounts = [
    ['U001', 'Guru Kelas 6 (Pengembang)', 'guru6@guru.sd.belajar.id', ROLE.SUPER, 'Kelas 6', '', 'L'],
    ['U002', 'Guru Kelas 1', 'guru1@guru.sd.belajar.id', ROLE.GURU, 'Kelas 1', '', 'P'],
    ['U003', 'Kepala Sekolah', 'kepsek@admin.sd.belajar.id', ROLE.KEPSEK, '', '', 'L'],
    ['U004', 'Guru PAI dan PJOK', 'mapel@guru.sd.belajar.id', ROLE.GURU, 'Kelas 1, Kelas 2, Kelas 3, Kelas 4, Kelas 5, Kelas 6', 'Pendidikan Agama, PJOK', 'P']
  ];
  const lines = ['Akun contoh (password sementara, tampil sekali):'];
  accounts.forEach(function (a) {
    const pw = randomPass_();
    lines.push('- ' + a[2] + '  |  ' + pw);
    appendUser_({ id_user: a[0], nama_lengkap: a[1], email: a[2], password_hash: hashPassword_(pw), role: a[3], kelas_ampu: a[4], mapel_ampu: a[5], id_folder_drive: '', status_aktif: 'aktif', jenis_kelamin: a[6] });
  });

  const nama6 = ['Ahmad Fauzan', 'Bunga Lestari', 'Citra Maharani', 'Dimas Prakoso', 'Eka Putri Ayu', 'Farhan Hidayat'];
  const nama1 = ['Gita Permata', 'Hana Safira', 'Indra Wijaya', 'Joko Santoso'];
  const siswa = [];
  nama6.forEach(function (n, i) { siswa.push(['S6' + pad2_(i + 1), '600' + pad2_(i + 1), n, 'Kelas 6', i % 2 ? 'P' : 'L', 'aktif']); });
  nama1.forEach(function (n, i) { siswa.push(['S1' + pad2_(i + 1), '100' + pad2_(i + 1), n, 'Kelas 1', i % 2 ? 'P' : 'L', 'aktif']); });
  appendRows_(SHEET.SISWA, siswa);

  const slot = [['1-2', '07.00', '08.10'], ['3-4', '08.30', '09.40'], ['5-6', '10.00', '11.10']];
  const mapelHari = {
    Senin: ['Pendidikan Pancasila', 'Bahasa Indonesia', 'Matematika'], Selasa: ['Matematika', 'IPAS', 'Bahasa Inggris'],
    Rabu: ['Bahasa Indonesia', 'Matematika', 'Seni Budaya'], Kamis: ['IPAS', 'Pendidikan Agama', 'Matematika'],
    Jumat: ['PJOK', 'Bahasa Indonesia', 'Pendidikan Agama'], Sabtu: ['Seni Budaya', 'IPAS', 'Muatan Lokal']
  };
  const guruMapel = { 'Pendidikan Agama': 'U004', 'PJOK': 'U004' };       // contoh guru mapel lintas kelas
  const jadwal = [];
  Object.keys(mapelHari).forEach(function (h) {
    mapelHari[h].forEach(function (m, i) { jadwal.push(['J' + h.slice(0, 2).toUpperCase() + (i + 1), h, 'Kelas 6', slot[i][0], slot[i][1], slot[i][2], m, guruMapel[m] || 'U001', 'Ruang Kelas 6']); });
  });
  appendRows_(SHEET.JADWAL, jadwal);

  const ph = '[CONTOH - ganti dengan teks resmi dari Kepmen/Capaian Pembelajaran]';
  appendRows_(SHEET.KURIKULUM, [
    ['TP-MTK6-01', 'Fase C', 'Kelas 6', 'Matematika', 'Bilangan', ph, '[CONTOH] Tujuan pembelajaran 1 mata pelajaran Matematika Kelas 6', '1', '[CONTOH] Topik materi', ''],
    ['TP-BIN6-01', 'Fase C', 'Kelas 6', 'Bahasa Indonesia', 'Membaca dan Memirsa', ph, '[CONTOH] Tujuan pembelajaran 1 mata pelajaran Bahasa Indonesia Kelas 6', '1', '[CONTOH] Topik materi', ''],
    ['TP-PAI6-01', 'Fase C', 'Kelas 6', 'Pendidikan Agama', 'Contoh elemen', ph, '[CONTOH] Tujuan pembelajaran 1 mata pelajaran Pendidikan Agama Kelas 6', '1', '[CONTOH] Topik materi', ''],
    ['TP-PJK6-01', 'Fase C', 'Kelas 6', 'PJOK', 'Contoh elemen', ph, '[CONTOH] Tujuan pembelajaran 1 mata pelajaran PJOK Kelas 6', '1', '[CONTOH] Topik materi', '']
  ]);
  return lines.join('\n');
}

function setupKonfigurasi_() {
  const p = props_(), cur = p.getProperties(), add = {};
  const def = { SCHOOL_NAME: 'SDN Grogol 02', SCHOOL_ADDRESS: '', GEMINI_MODEL: DEFAULT_MODEL };
  Object.keys(def).forEach(function (k) { if (cur[k] === undefined) add[k] = def[k]; });
  if (cur.GEMINI_MODEL === 'gemini-flash-latest' || cur.GEMINI_MODEL === 'gemini-2.5-flash') add.GEMINI_MODEL = DEFAULT_MODEL;   // upgrade ke model bawaan terbaru
  if (Object.keys(add).length) p.setProperties(add);
}

/** Membuat folder Drive per guru (di dalam folder induk) dan memberi akses lihat ke Kepala Sekolah. */
function siapkanFolderGuru() {
  const root = getRootFolder_();
  const tbl = readTable_(SHEET.USER);
  const sh = sheet_(SHEET.USER);
  const col = colIndex_(tbl.headers, 'id_folder_drive');
  tbl.rows.forEach(function (r) {
    const role = str_(r.role);
    if (role === ROLE.KEPSEK) {
      try { root.addViewer(str_(r.email)); } catch (e) { Logger.log('Tidak dapat membagikan folder ke ' + r.email + ': ' + e.message); }
      return;
    }
    let ok = false;
    if (str_(r.id_folder_drive)) { try { DriveApp.getFolderById(str_(r.id_folder_drive)); ok = true; } catch (e) { /* buat ulang */ } }
    if (!ok) {
      const f = root.createFolder(str_(r.nama_lengkap) + ' (' + str_(r.id_user) + ')');
      sh.getRange(r._row, col).setValue(f.getId());
    }
  });
  return root.getUrl();
}

/**
 * Akun dummy untuk uji coba (Super Admin + Guru Kelas 6, jadi semua fitur bisa dicoba).
 * Aman dijalankan berulang: jika email sudah ada, hanya password-nya yang disetel ulang.
 * HAPUS atau nonaktifkan akun ini di Master_User sebelum aplikasi dipakai guru sungguhan.
 */
function tambahAkunDummy() {
  setupSheet_(getSS_(), 'Master_User');            // pastikan kolom jenis_kelamin ada (migrasi v2.5)
  const email = 'ekowahyudi46@guru.sd.belajar.id', pw = 'Asdiru123';
  const tbl = readTable_(SHEET.USER);
  const ada = tbl.rows.filter(function (r) { return str_(r.email).toLowerCase() === email; })[0];
  if (ada) {
    aturPassword(email, pw);
    if (!str_(ada.jenis_kelamin)) sheet_(SHEET.USER).getRange(ada._row, colIndex_(tbl.headers, 'jenis_kelamin')).setValue('L');
  } else {
    appendUser_({ id_user: 'U900', nama_lengkap: 'Eko Wahyudi (Dummy)', email: email, password_hash: hashPassword_(pw), role: ROLE.SUPER, kelas_ampu: 'Kelas 6', mapel_ampu: '', id_folder_drive: '', status_aktif: 'aktif', jenis_kelamin: 'L' });
  }
  siapkanFolderGuru();
  Logger.log('Akun dummy siap: ' + email);
}

/** Jalankan dari editor: aturPassword('email@contoh.com', 'PasswordBaru123') */
function aturPassword(email, passwordBaru) {
  email = str_(email).toLowerCase();
  if (String(passwordBaru || '').length < 8) throw new Error('Password minimal 8 karakter.');
  const tbl = readTable_(SHEET.USER);
  const row = tbl.rows.filter(function (r) { return str_(r.email).toLowerCase() === email; })[0];
  if (!row) throw new Error('Email tidak ditemukan di Master_User.');
  sheet_(SHEET.USER).getRange(row._row, colIndex_(tbl.headers, 'password_hash')).setValue(hashPassword_(passwordBaru));
  CacheService.getScriptCache().remove('fail:' + email);          // buka kunci jika akun sempat terblokir 10 menit
  Logger.log('Password diperbarui untuk ' + email);
}

function menuAturPassword() {
  const ui = SpreadsheetApp.getUi();
  const e = ui.prompt('Atur password', 'Email pengguna:', ui.ButtonSet.OK_CANCEL);
  if (e.getSelectedButton() !== ui.Button.OK) return;
  const p = ui.prompt('Atur password', 'Password baru (minimal 8 karakter):', ui.ButtonSet.OK_CANCEL);
  if (p.getSelectedButton() !== ui.Button.OK) return;
  try { aturPassword(e.getResponseText(), p.getResponseText()); ui.alert('Password berhasil diperbarui.'); }
  catch (err) { ui.alert(err.message); }
}

/* =================== FOLDER DRIVE: BUKA DAN ATUR ALAMAT =================== */

/** Mengambil ID folder dari alamat Drive (…/folders/ID atau …?id=ID) atau dari ID polos. Kosong bila tidak dikenali. */
function folderIdDari_(teks) {
  const t = String(teks || '').trim();
  const m = t.match(/\/folders\/([A-Za-z0-9_-]{10,})/) || t.match(/[?&]id=([A-Za-z0-9_-]{10,})/);
  if (m) return m[1];
  return /^[A-Za-z0-9_-]{15,}$/.test(t) ? t : '';
}

/** Dialog berisi tautan ke folder induk arsip dan folder tiap pengguna. */
function menuBukaFolder() {
  const ui = SpreadsheetApp.getUi();
  try {
    const root = getRootFolder_(), tbl = readTable_(SHEET.USER);
    const h = function (x) { return String(x).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); };
    const url = function (id) { return 'https://drive.google.com/drive/folders/' + id; };
    const baris = [];
    const fi = root.getFoldersByName('Foto Profil');
    if (fi.hasNext()) baris.push(['Foto Profil', fi.next().getId()]);
    tbl.rows.forEach(function (r) {
      const id = str_(r.id_folder_drive);
      if (/^[A-Za-z0-9_-]{10,}$/.test(id)) baris.push([str_(r.nama_lengkap), id]);
    });
    const tr = baris.map(function (b) {
      return '<tr><td>' + h(b[0]) + '</td><td class="r"><a href="' + url(b[1]) + '" target="_blank" rel="noopener">Buka</a></td></tr>';
    }).join('');
    const html = HtmlService.createHtmlOutput(
      '<style>body{font:14px Arial,sans-serif;margin:14px;color:#222}h3{margin:0 0 4px;font-size:15px}.nm{color:#555;margin-bottom:10px;word-break:break-all}' +
      '.btn{display:inline-block;background:#0A5BD8;color:#fff;text-decoration:none;padding:10px 18px;border-radius:6px;font-weight:bold}' +
      '.url{margin:10px 0 16px;font-size:12px;color:#666;word-break:break-all}table{width:100%;border-collapse:collapse}td{padding:7px 4px;border-bottom:1px solid #eee}' +
      '.r{text-align:right}a{color:#0A5BD8;font-weight:bold}.box{max-height:210px;overflow:auto;border:1px solid #eee;border-radius:6px;padding:0 8px}.hint{color:#777;font-size:12px;margin-top:10px}</style>' +
      '<h3>Folder induk arsip</h3><div class="nm">' + h(root.getName()) + '</div>' +
      '<a class="btn" href="' + url(root.getId()) + '" target="_blank" rel="noopener">Buka folder induk di Drive</a>' +
      '<div class="url">' + url(root.getId()) + '</div>' +
      '<h3>Folder per pengguna</h3>' +
      (tr ? '<div class="box"><table>' + tr + '</table></div>' : '<div class="hint">Belum ada folder pengguna. Jalankan "Siapkan folder Drive guru".</div>') +
      '<div class="hint">Untuk memindahkan lokasi folder induk, gunakan menu "Atur alamat folder Drive".</div>'
    ).setWidth(460).setHeight(480);
    ui.showModalDialog(html, 'Folder Drive ASDIRU');
  } catch (err) { ui.alert('ASDIRU', err.message, ui.ButtonSet.OK); }
}

/** Menyetel folder Drive yang sudah ada (dari alamat/ID) sebagai folder induk arsip. */
function menuAturFolderDrive() {
  const ui = SpreadsheetApp.getUi();
  try {
    requireOwner_();
    const p = props_(), idLama = p.getProperty('ROOT_FOLDER_ID') || '';
    const r = ui.prompt('Atur alamat folder Drive',
      'Tempel alamat (URL) atau ID folder Google Drive yang akan menjadi folder induk arsip.\n' +
      'Contoh: https://drive.google.com/drive/folders/1AbCdEfGh...\n\nAkun Anda harus punya akses edit ke folder itu.', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    const id = folderIdDari_(r.getResponseText());
    if (!id) fail_('Alamat atau ID folder tidak dikenali. Buka folder di Drive lalu salin alamat dari kolom alamat browser.');
    if (id === idLama) { ui.alert('ASDIRU', 'Folder induk sudah memakai folder ini.', ui.ButtonSet.OK); return; }
    let baru;
    try { baru = DriveApp.getFolderById(id); } catch (e) { fail_('Folder tidak dapat dibuka. Pastikan alamatnya benar dan akun Anda punya akses.'); }
    try { baru.createFolder('_uji_akses').setTrashed(true); } catch (e) { fail_('Akun Anda hanya punya akses lihat. Dibutuhkan akses edit pada folder tersebut.'); }

    const pindah = ui.alert('Pindahkan folder yang sudah ada?',
      'Folder induk akan diganti ke:\n"' + baru.getName() + '"\n\n' +
      'YA: pindahkan folder guru dan folder Foto Profil yang sudah ada ke folder ini.\n' +
      'TIDAK: folder lama tetap di tempatnya, hanya folder baru yang dibuat di sini.\n' +
      'BATAL: tidak mengubah apa pun.', ui.ButtonSet.YES_NO_CANCEL);
    if (pindah === ui.Button.CANCEL || pindah === ui.Button.CLOSE) return;

    p.setProperty('ROOT_FOLDER_ID', id);
    let n = 0;
    if (pindah === ui.Button.YES) {
      readTable_(SHEET.USER).rows.forEach(function (row) {
        const fid = str_(row.id_folder_drive);
        if (!fid) return;
        try { DriveApp.getFolderById(fid).moveTo(baru); n++; } catch (e) { Logger.log('Gagal memindahkan folder ' + fid + ': ' + e.message); }
      });
      if (idLama) {
        try { const it = DriveApp.getFolderById(idLama).getFoldersByName('Foto Profil'); if (it.hasNext()) it.next().moveTo(baru); } catch (e) { /* folder lama tidak ada */ }
      }
    }
    siapkanFolderGuru();                                       // buat folder yang belum ada dan bagikan ke Kepala Sekolah
    ui.alert('Selesai', 'Folder induk sekarang: "' + baru.getName() + '"' + (pindah === ui.Button.YES ? '\n' + n + ' folder pengguna dipindahkan.' : '') +
      '\n\n' + baru.getUrl(), ui.ButtonSet.OK);
  } catch (err) { ui.alert('ASDIRU', err.message, ui.ButtonSet.OK); }
}

/* =================== MENU KELOLA PENGGUNA & API KEY (pemilik spreadsheet) =================== */

/** Menu hanya untuk pemilik spreadsheet. Pemilik tidak terdeteksi (mis. Shared Drive) = diizinkan. */
function requireOwner_() {
  let owner = null, me = '';
  try { owner = getSS_().getOwner(); } catch (e) { /* tidak terbaca */ }
  try { me = Session.getEffectiveUser().getEmail(); } catch (e) { /* tidak terbaca */ }
  if (owner && owner.getEmail() && me && owner.getEmail().toLowerCase() !== me.toLowerCase()) {
    fail_('Menu ini hanya untuk pemilik spreadsheet (' + owner.getEmail() + ').');
  }
}

function uiAlert_(judul, pesan) { SpreadsheetApp.getUi().alert(judul, pesan, SpreadsheetApp.getUi().ButtonSet.OK); }

function jumlahSuperAktif_(kecualiRow) {
  return rows_(SHEET.USER).filter(function (r) {
    return str_(r.role) === ROLE.SUPER && str_(r.status_aktif).toLowerCase() !== 'nonaktif' && r._row !== kecualiRow;
  }).length;
}

function cariUserByEmail_(email) {
  email = str_(email).toLowerCase();
  const tbl = readTable_(SHEET.USER);
  const row = tbl.rows.filter(function (r) { return str_(r.email).toLowerCase() === email; })[0];
  if (!row) fail_('Email "' + email + '" tidak ditemukan di Master_User.');
  return { tbl: tbl, row: row };
}

function cabutSesi_(idUser) {
  CacheService.getScriptCache().put('revoked:' + str_(idUser), String(Date.now()), SESSION_TTL);   // sesi lama langsung tidak berlaku
}

/** Formulir tambah pengguna (nama, email, role, kelas ampu, password). */
function menuTambahPengguna() {
  try { requireOwner_(); } catch (e) { uiAlert_('ASDIRU', e.message); return; }
  const kelas = KELAS_ALL.map(function (k) {
    return '<label class="cb"><input type="checkbox" name="kelas" value="' + k + '"> ' + k + '</label>';
  }).join('');
  const html = HtmlService.createHtmlOutput(
    '<style>body{font:14px Arial,sans-serif;margin:12px;color:#222}label{display:block;margin:8px 0 3px;font-weight:bold}' +
    'input[type=text],input[type=email],input[type=password],select{width:100%;box-sizing:border-box;padding:7px}' +
    '.cb{display:inline-block;width:32%;font-weight:normal;margin:2px 0}.hint{color:#666;font-size:12px}' +
    'button{margin-top:14px;padding:9px 18px;background:#0A5BD8;color:#fff;border:0;border-radius:4px;font-size:14px;cursor:pointer}' +
    'button:disabled{opacity:.6}#msg{margin-top:10px;font-weight:bold}.err{color:#B00020}.ok{color:#0B7A2B}</style>' +
    '<label>Nama lengkap</label><input type="text" id="nama">' +
    '<label>Email</label><input type="email" id="email" placeholder="nama@email.com">' +
    '<label>Jenis kelamin (untuk sapaan Pak/Bu)</label><select id="g"><option value="">- pilih -</option><option value="L">Laki-laki (Pak)</option><option value="P">Perempuan (Bu)</option></select>' +
    '<label>Role</label><select id="role"><option value="guru">Guru</option><option value="kepala_sekolah">Kepala Sekolah</option><option value="super_admin">Super Admin</option></select>' +
    '<label>Kelas ampu (wajib untuk Guru)</label><div>' + kelas + '</div>' +
    '<label>Mapel ampu (guru mapel, pisahkan koma)</label><input type="text" id="mapel" placeholder="Pendidikan Agama, PJOK"><div class="hint">Isi persis seperti nama mapel di Master_Kurikulum.</div>' +
    '<label>Password awal</label><input type="password" id="pw" autocomplete="new-password"><div class="hint">Minimal 8 karakter.</div>' +
    '<button id="b" onclick="kirim()">Simpan pengguna</button><div id="msg"></div>' +
    '<script>function $(i){return document.getElementById(i)}' +
    'function kirim(){var k=[].slice.call(document.querySelectorAll("input[name=kelas]:checked")).map(function(x){return x.value});' +
    'var u={nama:$("nama").value,email:$("email").value,gender:$("g").value,role:$("role").value,kelas:k,mapel:$("mapel").value,password:$("pw").value};' +
    '$("b").disabled=true;$("msg").className="";$("msg").textContent="Menyimpan...";' +
    'google.script.run.withSuccessHandler(function(m){$("msg").className="ok";$("msg").textContent=m;$("pw").value="";$("b").disabled=false})' +
    '.withFailureHandler(function(e){$("msg").className="err";$("msg").textContent=e.message;$("b").disabled=false}).menuSimpanPengguna(u)}</script>'
  ).setWidth(480).setHeight(640);
  SpreadsheetApp.getUi().showModalDialog(html, 'Tambah pengguna ASDIRU');
}

/** Dipanggil dari formulir di atas (harus publik agar bisa dipanggil google.script.run). */
function menuSimpanPengguna(u) {
  requireOwner_();
  u = u || {};
  const nama = cleanText_(u.nama, 100), email = str_(u.email).toLowerCase(), role = str_(u.role), g = str_(u.gender).toUpperCase();
  const pw = String(u.password || '');
  if (!nama) fail_('Nama lengkap wajib diisi.');
  if (!EMAIL_RE.test(email)) fail_('Masukkan alamat email yang valid.');
  if ([ROLE.SUPER, ROLE.GURU, ROLE.KEPSEK].indexOf(role) < 0) fail_('Role tidak valid.');
  if (g !== 'L' && g !== 'P') fail_('Pilih jenis kelamin agar sapaan di dasbor sesuai.');
  if (pw.length < 8) fail_('Password minimal 8 karakter.');
  const kelas = (Array.isArray(u.kelas) ? u.kelas : []).map(str_).filter(function (k) { return KELAS_ALL.indexOf(k) >= 0; });
  if (role === ROLE.GURU && !kelas.length) fail_('Guru harus memiliki minimal satu kelas ampu.');
  const mapel = cleanText_(u.mapel, 200);
  return withLock_(function () {
    const rows = rows_(SHEET.USER);
    if (rows.some(function (r) { return str_(r.email).toLowerCase() === email; })) fail_('Email sudah terdaftar.');
    let max = 0;
    rows.forEach(function (r) { const m = str_(r.id_user).match(/^U(\d+)$/); if (m) max = Math.max(max, Number(m[1])); });
    const id = 'U' + ('00' + (max + 1)).slice(-3);
    appendUser_({ id_user: id, nama_lengkap: nama, email: email, password_hash: hashPassword_(pw), role: role, kelas_ampu: kelas.join(', '), mapel_ampu: mapel, id_folder_drive: '', status_aktif: 'aktif', jenis_kelamin: g });
    if (role === ROLE.KEPSEK) { try { getRootFolder_().addViewer(email); } catch (e) { /* domain membatasi berbagi */ } }
    else ensureFolderGuru_(id);
    return 'Pengguna ' + email + ' ditambahkan (ID ' + id + ').';
  });
}

/** Menonaktifkan akun (data tetap ada, login ditolak) atau mengaktifkannya kembali. */
function menuToggleStatus() {
  const ui = SpreadsheetApp.getUi();
  try {
    requireOwner_();
    const e = ui.prompt('Nonaktifkan / aktifkan pengguna', 'Email pengguna:', ui.ButtonSet.OK_CANCEL);
    if (e.getSelectedButton() !== ui.Button.OK) return;
    const f = cariUserByEmail_(e.getResponseText());
    const aktifSekarang = str_(f.row.status_aktif).toLowerCase() !== 'nonaktif';
    if (aktifSekarang && str_(f.row.role) === ROLE.SUPER && jumlahSuperAktif_(f.row._row) < 1) fail_('Tidak bisa menonaktifkan Super Admin aktif terakhir.');
    const yakin = ui.alert('Konfirmasi', (aktifSekarang ? 'Nonaktifkan ' : 'Aktifkan kembali ') + str_(f.row.nama_lengkap) + ' (' + str_(f.row.email) + ')?', ui.ButtonSet.YES_NO);
    if (yakin !== ui.Button.YES) return;
    sheet_(SHEET.USER).getRange(f.row._row, colIndex_(f.tbl.headers, 'status_aktif')).setValue(aktifSekarang ? 'nonaktif' : 'aktif');
    if (aktifSekarang) cabutSesi_(f.row.id_user);
    ui.alert('Selesai', 'Akun ' + str_(f.row.email) + (aktifSekarang ? ' dinonaktifkan.' : ' diaktifkan kembali.'), ui.ButtonSet.OK);
  } catch (err) { ui.alert('ASDIRU', err.message, ui.ButtonSet.OK); }
}

/** Menghapus baris pengguna secara permanen. Folder Drive dan arsip lama tidak ikut terhapus. */
function menuHapusPengguna() {
  const ui = SpreadsheetApp.getUi();
  try {
    requireOwner_();
    const e = ui.prompt('Hapus pengguna', 'Email pengguna yang akan dihapus permanen:', ui.ButtonSet.OK_CANCEL);
    if (e.getSelectedButton() !== ui.Button.OK) return;
    const f = cariUserByEmail_(e.getResponseText());
    if (str_(f.row.role) === ROLE.SUPER && jumlahSuperAktif_(f.row._row) < 1) fail_('Tidak bisa menghapus Super Admin aktif terakhir.');
    const yakin = ui.alert('Hapus permanen?',
      str_(f.row.nama_lengkap) + ' (' + str_(f.row.email) + ') akan dihapus dari Master_User dan tidak bisa login lagi.\n\n' +
      'Folder Drive dan arsip miliknya tidak dihapus. Jika hanya ingin menutup akses sementara, gunakan "Nonaktifkan / aktifkan pengguna".',
      ui.ButtonSet.YES_NO);
    if (yakin !== ui.Button.YES) return;
    cabutSesi_(f.row.id_user);
    sheet_(SHEET.USER).deleteRow(f.row._row);
    ui.alert('Selesai', 'Pengguna ' + str_(f.row.email) + ' dihapus.', ui.ButtonSet.OK);
  } catch (err) { ui.alert('ASDIRU', err.message, ui.ButtonSet.OK); }
}

/**
 * Pemilik boleh mengetik password polos langsung di kolom password_hash (mis. saat menambah baris manual).
 * Menu ini mengubah semuanya menjadi hash sekarang juga, tanpa menunggu login pertama.
 */
function amankanPasswordPolos() {
  const ui = SpreadsheetApp.getUi();
  try {
    requireOwner_();
    const tbl = readTable_(SHEET.USER), sh = sheet_(SHEET.USER), col = colIndex_(tbl.headers, 'password_hash');
    let n = 0, pendek = [];
    tbl.rows.forEach(function (r) {
      const v = str_(r.password_hash);
      if (!v || v.indexOf('v1$') === 0) return;
      if (v.length < 8) { pendek.push(str_(r.email)); return; }
      sh.getRange(r._row, col).setValue(hashPassword_(v));
      n++;
    });
    ui.alert('Selesai', n + ' password diamankan.' + (pendek.length ? '\n\nDilewati (kurang dari 8 karakter): ' + pendek.join(', ') : ''), ui.ButtonSet.OK);
  } catch (err) { ui.alert('ASDIRU', err.message, ui.ButtonSet.OK); }
}

/** Dialog API Key dengan kolom tersamarkan (password). Kunci disimpan di Script Properties, tidak di sheet. */
function menuApiKey() {
  try { requireOwner_(); } catch (e) { uiAlert_('ASDIRU', e.message); return; }
  const terpasang = !!props_().getProperty('GEMINI_API_KEY');
  const html = HtmlService.createHtmlOutput(
    '<style>body{font:14px Arial,sans-serif;margin:12px;color:#222}input{width:100%;box-sizing:border-box;padding:8px;margin-top:6px}' +
    'button{margin-top:14px;padding:9px 18px;background:#0A5BD8;color:#fff;border:0;border-radius:4px;font-size:14px;cursor:pointer}' +
    'button:disabled{opacity:.6}.hint{color:#666;font-size:12px;margin-top:6px}#msg{margin-top:10px;font-weight:bold}.err{color:#B00020}.ok{color:#0B7A2B}</style>' +
    '<div>' + (terpasang ? 'API Key sudah terpasang. Isi di bawah untuk menggantinya.' : 'API Key belum terpasang.') + '</div>' +
    '<input type="password" id="k" autocomplete="off" placeholder="Tempel API Key Gemini">' +
    '<div class="hint">Dapatkan dari Google AI Studio. Kunci disimpan di server dan tidak pernah ditampilkan kembali.</div>' +
    '<button id="b" onclick="kirim()">Simpan API Key</button><div id="msg"></div>' +
    '<script>function $(i){return document.getElementById(i)}' +
    'function kirim(){$("b").disabled=true;$("msg").className="";$("msg").textContent="Menyimpan...";' +
    'google.script.run.withSuccessHandler(function(m){$("msg").className="ok";$("msg").textContent=m;$("k").value="";$("b").disabled=false})' +
    '.withFailureHandler(function(e){$("msg").className="err";$("msg").textContent=e.message;$("b").disabled=false}).menuSimpanApiKey($("k").value)}</script>'
  ).setWidth(420).setHeight(240);
  SpreadsheetApp.getUi().showModalDialog(html, 'API Key Gemini');
}

function menuSimpanApiKey(key) {
  requireOwner_();
  key = str_(key);
  if (key.length < 20 || /\s/.test(key)) fail_('API Key tidak valid. Salin ulang dari Google AI Studio.');
  props_().setProperty('GEMINI_API_KEY', key);
  return 'API Key tersimpan.';
}

/** Mengatur model Gemini ke model bawaan aplikasi (DEFAULT_MODEL di code.gs) tanpa membuka Panel Admin. */
function perbaruiModelGemini() {
  try { requireOwner_(); } catch (e) { try { uiAlert_('ASDIRU', e.message); } catch (x) { Logger.log(e.message); } return; }
  const p = props_();
  p.setProperty('GEMINI_MODEL', DEFAULT_MODEL);
  p.deleteProperty('GEMINI_MODEL_LAST_USED');
  const msg = 'Model Gemini diatur ke ' + DEFAULT_MODEL + '.\n\nUji dengan tombol uji koneksi di Panel Admin (Mesin AI).';
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('ASDIRU', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { /* lihat Logger */ }
}

function cekInstalasi() {
  const out = [];
  const ss = getSS_();
  Object.keys(SETUP_SCHEMA).forEach(function (n) {
    const sh = ss.getSheetByName(n);
    out.push((sh ? 'OK   ' : 'HILANG ') + n + (sh ? ' (' + Math.max(sh.getLastRow() - 1, 0) + ' baris data)' : ''));
  });
  const p = props_();
  out.push(p.getProperty('GEMINI_API_KEY') ? 'OK   API Key Gemini terisi' : 'BELUM API Key Gemini (isi di Panel Admin)');
  out.push('Model Gemini: ' + (p.getProperty('GEMINI_MODEL') || DEFAULT_MODEL));
  out.push(pedomanInfo_('PROMPT_MODUL').aktif ? 'OK   Pedoman modul aktif' : 'BELUM Pedoman modul');
  out.push(pedomanInfo_('PROMPT_ASESMEN').aktif ? 'OK   Pedoman asesmen aktif' : 'BELUM Pedoman asesmen');
  out.push(p.getProperty('ROOT_FOLDER_ID') ? 'OK   Folder induk Drive siap' : 'BELUM Folder induk Drive');
  const msg = out.join('\n');
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert('Cek instalasi ASDIRU', msg, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { /* lihat Logger */ }
  return msg;
}