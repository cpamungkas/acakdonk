'use strict';
const KEY = 'luckydraw_v1';
const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const norm = v => String(v ?? '').replace(/\s+/g, '').toUpperCase();
const txt = v => String(v ?? '').replace(/\s+/g, ' ').trim();

let S = load();
function load() {
  try { const d = JSON.parse(localStorage.getItem(KEY)); if (d && Array.isArray(d.peserta) && Array.isArray(d.hadiah) && Array.isArray(d.winners)) return d; } catch (e) {}
  return { peserta: [], hadiah: [], winners: [] };
}
function save(quiet) {
  try { localStorage.setItem(KEY, JSON.stringify(S)); if (!quiet) toast('Data berhasil disimpan'); }
  catch (e) { Swal.fire('Gagal menyimpan', 'Penyimpanan browser penuh atau diblokir.', 'error'); }
}
const toast = t => Swal.fire({ toast: true, position: 'top-end', icon: 'success', title: t, showConfirmButton: false, timer: 1400 });

/* ---------- logika inti ---------- */
const eligible = (p, h) => h.rule === 'ALL' || p.lokasi === 'PULAUJAWA';
const candidates = h => { const used = new Set(S.winners.map(w => w.nik)); return S.peserta.filter(p => !used.has(p.nik) && eligible(p, h)); };
const given = id => S.winners.filter(w => w.hadiahId === id && w.status === 'HADIR').length;
const remaining = h => h.qty - given(h.id);
const pending = () => S.winners.find(w => w.status === 'PENDING');
function rnd(n) { // uniform, tanpa modulo bias
  const a = new Uint32Array(1), lim = Math.floor(4294967296 / n) * n;
  do crypto.getRandomValues(a); while (a[0] >= lim);
  return a[0] % n;
}

/* ---------- import ---------- */
function readFile(f) {
  return new Promise((ok, no) => {
    const r = new FileReader();
    r.onload = e => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: false, raw: true });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { raw: false, defval: '' });
        ok(rows.map(r => { const o = {}; for (const k in r) o[k.trim().toUpperCase()] = r[k]; return o; }));
      } catch (x) { no(x); }
    };
    r.onerror = no; r.readAsArrayBuffer(f);
  });
}
async function doImport(type, file) {
  if (S.winners.length) return Swal.fire('Tidak dapat import', 'Undian sudah berjalan. Reset data terlebih dahulu untuk mengganti master.', 'warning');
  let rows;
  try { rows = await readFile(file); } catch (e) { return Swal.fire('Gagal', 'File tidak dapat dibaca.', 'error'); }
  const cols = type === 'peserta' ? ['NIK', 'NAMA', 'DIVISI', 'UNIT', 'LOKASI'] : ['NAMAHADIAH', 'QTYHADIAH', 'RULE'];
  if (!rows.length || cols.some(c => !(c in rows[0]))) return Swal.fire('Header tidak sesuai', 'Kolom wajib: ' + cols.join(', '), 'error');
  const ok = [], err = [], seen = new Set();
  rows.forEach((r, i) => {
    const n = i + 2, e = [];
    if (type === 'peserta') {
      const p = { nik: txt(r.NIK), nama: txt(r.NAMA), divisi: txt(r.DIVISI), unit: txt(r.UNIT), lokasi: norm(r.LOKASI) };
      if (!p.nik) e.push('NIK kosong'); if (!p.nama) e.push('Nama kosong'); if (!p.divisi) e.push('DIVISI kosong'); if (!p.unit) e.push('UNIT kosong');
      if (!p.lokasi) e.push('LOKASI kosong'); else if (!['PULAUJAWA', 'LUARPULAUJAWA'].includes(p.lokasi)) e.push('LOKASI tidak valid (' + txt(r.LOKASI) + ')');
      if (p.nik && seen.has(p.nik)) e.push('NIK duplikat'); 
      e.length ? err.push([n, e.join('; ')]) : (seen.add(p.nik), ok.push(p));
    } else {
      const h = { nama: txt(r.NAMAHADIAH), qty: txt(r.QTYHADIAH), rule: norm(r.RULE) };
      if (!h.nama) e.push('NamaHadiah kosong');
      if (!/^\d+$/.test(h.qty) || +h.qty < 1) e.push('QtyHadiah harus integer > 0'); else h.qty = +h.qty;
      if (!h.rule) e.push('Rule kosong'); else if (!['ALL', 'PULAUJAWA'].includes(h.rule)) e.push('Rule tidak valid (' + txt(r.RULE) + ')');
      e.length ? err.push([n, e.join('; ')]) : ok.push(h);
    }
  });
  const head = cols.map(c => `<th>${c}</th>`).join('');
  const body = ok.map(o => '<tr>' + Object.values(o).map(v => `<td>${esc(v)}</td>`).join('') + '</tr>').join('');
  const res = await Swal.fire({
    title: 'Preview Import ' + (type === 'peserta' ? 'Peserta' : 'Hadiah'), width: 900, showCancelButton: true,
    showConfirmButton: ok.length > 0, confirmButtonText: `Import ${ok.length} data valid`, cancelButtonText: 'Batal',
    html: `<p>Total baris: <b>${rows.length}</b> | Valid: <b style="color:#2fa863">${ok.length}</b> | Invalid: <b style="color:#d24a4a">${err.length}</b></p>
      ${err.length ? `<div style="max-height:130px;overflow:auto;text-align:left;font-size:.85rem;color:#d24a4a">${err.map(x => `Baris ${x[0]}: ${esc(x[1])}`).join('<br>')}<br><i>Baris invalid tidak akan diimport.</i></div>` : ''}
      <div style="max-height:260px;overflow:auto;margin-top:8px"><table style="width:100%;font-size:.8rem"><tr>${head}</tr>${body}</table></div>`
  });
  if (!res.isConfirmed) return;
  if (type === 'peserta') S.peserta = ok; else S.hadiah = ok.map((h, i) => ({ id: 'H' + String(i + 1).padStart(3, '0'), ...h }));
  save(true); render();
  Swal.fire('Berhasil', `Import master ${type} berhasil (${ok.length} data, ${err.length} ditolak).`, 'success');
}

/* ---------- render ---------- */
const table = (cols, rows) => rows.length ? `<table><tr>${cols.map(c => `<th>${c}</th>`).join('')}</tr>${rows.map(r => '<tr>' + r.map(v => `<td>${esc(v)}</td>`).join('') + '</tr>').join('')}</table>` : '<div class="empty">Belum ada data.</div>';
function render() { renderDash(); renderMaster(); renderDraw(); renderWin(); }
function renderDash() {
  const hadir = S.winners.filter(w => w.status === 'HADIR').length, absen = S.winners.filter(w => w.status === 'TIDAK HADIR').length;
  const tot = S.hadiah.reduce((a, h) => a + h.qty, 0);
  $('#cards').innerHTML = [['Total peserta', S.peserta.length], ['Total hadiah', tot], ['Hadiah sudah diundi', hadir], ['Hadiah belum diundi', tot - hadir], ['Pemenang hadir', hadir], ['Pemenang tidak hadir', absen]]
    .map(([a, b]) => `<div class="card"><span>${a}</span><b>${b}</b></div>`).join('');
  const b = (t, y) => `<span class="badge ${y ? 'y' : 'n'}">${t}: ${y ? 'Siap' : 'Belum diimport'}</span>`;
  $('#status').innerHTML = b('Master peserta', S.peserta.length) + b('Master hadiah', S.hadiah.length) + b('Kesiapan undian', S.peserta.length && S.hadiah.length);
}
function renderMaster() {
  $('#tPeserta').innerHTML = table(['NIK', 'Nama', 'Divisi', 'Unit'], S.peserta.map(p => [p.nik, p.nama, p.divisi, p.unit]));
  $('#tHadiah').innerHTML = table(['ID', 'Nama Hadiah', 'Qty', 'Sudah diberikan', 'Sisa'], S.hadiah.map(h => [h.id, h.nama, h.qty, given(h.id), remaining(h)]));
}

/* ---------- undian ---------- */
let tm = null;
const pendings = () => S.winners.filter(w => w.status === 'PENDING');
const scale = n => (n <= 2 ? 2.4 : n <= 5 ? 1.6 : n <= 10 ? 1.15 : .9) + 'rem';
function renderDraw() {
  const sel = $('#selH'), cur = sel.value, pd = pendings();
  sel.innerHTML = '<option value="">-- Pilih hadiah --</option>' + S.hadiah.map(h => `<option value="${h.id}">${esc(h.nama)} (sisa ${remaining(h)}/${h.qty})</option>`).join('');
  sel.value = pd.length ? pd[0].hadiahId : cur;
  if (tm) return;
  const h = S.hadiah.find(x => x.id === sel.value);
  $('#prize').textContent = h ? h.nama : '';
  $('#hInfo').textContent = h ? `Tersedia ${remaining(h)} dari ${h.qty} | Ketentuan: ${h.rule === 'ALL' ? 'Semua peserta' : 'Hanya Pulau Jawa'} | Kandidat: ${candidates(h).length}` : '';
  if (pd.length) {
    $('#stage').style.setProperty('--fs', scale(pd.length));
    $('#slots').className = 'slots';
    $('#slots').innerHTML = pd.map(w => `<div class="wrow"><span class="nik">${esc(w.nik)}</span><b class="nm">${esc(w.nama)}</b><span class="un">${esc(w.unit)} - ${esc(w.divisi)}</span><span class="act"><button class="ok" data-n="${esc(w.nik)}" data-s="HADIR">Hadir</button><button class="bad" data-n="${esc(w.nik)}" data-s="TIDAK HADIR">Tidak Hadir</button></span></div>`).join('');
    $('#stageMsg').textContent = `${pd.length} pemenang menunggu konfirmasi kehadiran`;
    $('#btnStart').hidden = true; $('#btnStop').hidden = true; sel.disabled = true;
    return;
  }
  $('#slots').innerHTML = '';
  $('#stageMsg').textContent = h ? (remaining(h) > 0 ? `Siap mengundi ${remaining(h)} pemenang` : 'Hadiah ini sudah selesai diundi') : 'Pilih hadiah, lalu mulai';
  const redo = h && S.winners.some(w => w.hadiahId === h.id && w.status === 'TIDAK HADIR');
  $('#btnStart').hidden = false; $('#btnStart').disabled = !h || remaining(h) <= 0;
  $('#btnStart').textContent = redo ? 'UNDI ULANG' : 'MULAI UNDIAN';
  $('#btnStop').hidden = true; sel.disabled = false;
}
function start() {
  if (!S.peserta.length || !S.hadiah.length) return Swal.fire('Master data belum lengkap', 'Import master peserta dan hadiah terlebih dahulu.', 'warning');
  if (pendings().length) return renderDraw();
  const h = S.hadiah.find(x => x.id === $('#selH').value);
  if (!h) return Swal.fire('Pilih hadiah', 'Pilih hadiah yang akan diundi.', 'info');
  if (remaining(h) <= 0) return Swal.fire('Hadiah selesai', 'Seluruh slot hadiah ini sudah diberikan.', 'info');
  const c = candidates(h);
  if (!c.length) return Swal.fire('Tidak ada kandidat', 'Tidak ada peserta yang memenuhi syarat untuk hadiah ini.', 'warning');
  const k = Math.min(remaining(h), c.length);
  $('#selH').disabled = true; $('#btnStart').hidden = true; $('#btnStop').hidden = false;
  $('#stage').style.setProperty('--fs', scale(k));
  $('#slots').className = 'slots grid';
  $('#slots').innerHTML = Array.from({ length: k }, () => '<div class="wcard"><span></span></div>').join('');
  $('#stageMsg').textContent = k < remaining(h) ? `Kandidat hanya cukup untuk ${k} dari ${remaining(h)} slot` : `Mengacak ${k} pemenang...`;
  tm = setInterval(() => document.querySelectorAll('#slots .wcard span').forEach(s => { s.textContent = c[rnd(c.length)].nama; }), 70);
  toast('Pengundian dimulai');
}
function stop() {
  clearInterval(tm); tm = null;
  const h = S.hadiah.find(x => x.id === $('#selH').value), c = candidates(h).slice(), k = Math.min(remaining(h), c.length);
  if (!k) { renderDraw(); return Swal.fire('Tidak ada kandidat', 'Tidak ada peserta yang memenuhi syarat.', 'warning'); }
  const now = new Date().toISOString();
  for (let i = 0; i < k; i++) { // pilih k peserta berbeda
    const j = i + rnd(c.length - i); [c[i], c[j]] = [c[j], c[i]];
    const p = c[i];
    S.winners.push({ nik: p.nik, nama: p.nama, divisi: p.divisi, unit: p.unit, lokasi: p.lokasi, hadiahId: h.id, hadiah: h.nama, status: 'PENDING', waktu: now });
  }
  save(true); renderDash(); renderDraw();
  toast(k + ' pemenang berhasil dipilih');
}
function mark(nik, status) {
  const w = S.winners.find(x => x.nik === nik && x.status === 'PENDING'); if (!w) return;
  w.status = status; save(); render();
  if (pendings().length) return;
  const h = S.hadiah.find(x => x.id === w.hadiahId), r = remaining(h);
  if (r > 0) Swal.fire('Konfirmasi selesai', `Masih ada ${r} slot hadiah. Tekan UNDI ULANG untuk memilih peserta baru.`, 'info');
  else Swal.fire('Selesai', 'Seluruh slot hadiah ini sudah diberikan.', 'success');
}

/* ---------- daftar pemenang ---------- */
function filtered() {
  const q = $('#fQ').value.toLowerCase(), h = $('#fH').value, s = $('#fS').value;
  return S.winners.filter(w => w.status !== 'PENDING' && (!q || w.nama.toLowerCase().includes(q) || w.nik.toLowerCase().includes(q)) && (!h || w.hadiahId === h) && (!s || w.status === s));
}
function renderWin() {
  const fh = $('#fH'), cur = fh.value;
  fh.innerHTML = '<option value="">Semua hadiah</option>' + S.hadiah.map(h => `<option value="${h.id}">${esc(h.nama)}</option>`).join(''); fh.value = cur;
  const f = filtered(), row = (w, i) => [i + 1, w.nik, w.nama, w.divisi, w.unit, w.lokasi, w.hadiah, w.status, new Date(w.waktu).toLocaleString('id-ID')];
  const cols = ['No.', 'NIK', 'Nama', 'Divisi', 'Unit', 'Lokasi', 'Hadiah', 'Status', 'Waktu'];
  const a = f.filter(w => w.status === 'HADIR'), b = f.filter(w => w.status === 'TIDAK HADIR');
  $('#tHadir').innerHTML = table(cols, a.map(row)); $('#tAbsen').innerHTML = table(cols, b.map(row));
  $('#nH').textContent = `(${a.length})`; $('#nA').textContent = `(${b.length})`;
}
function exportWin(fmt) {
  const f = filtered(); if (!f.length) return Swal.fire('Tidak ada data', 'Belum ada hasil undian untuk diexport.', 'info');
  const data = f.map((w, i) => ({ No: i + 1, NIK: w.nik, Nama: w.nama, Divisi: w.divisi, Unit: w.unit, Lokasi: w.lokasi, Hadiah: w.hadiah, Status: w.status, Waktu: new Date(w.waktu).toLocaleString('id-ID') }));
  if (fmt === 'pdf') {
    const d = new window.jspdf.jsPDF({ orientation: 'landscape' });
    d.setFontSize(14); d.text('Hasil Undian Hadiah', 14, 14);
    d.autoTable({ startY: 20, head: [Object.keys(data[0])], body: data.map(Object.values), styles: { fontSize: 8 }, headStyles: { fillColor: [10, 19, 48], textColor: [217, 178, 90] } });
    d.save('hasil-undian.pdf'); toast('Export berhasil'); return;
  }
  const ws = XLSX.utils.json_to_sheet(data);
  if (fmt === 'xlsx') { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Hasil Undian'); XLSX.writeFile(wb, 'hasil-undian.xlsx'); }
  else download('hasil-undian.csv', '\ufeff' + XLSX.utils.sheet_to_csv(ws), 'text/csv');
  toast('Export berhasil');
}
function download(name, content, type) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([content], { type })); a.download = name; a.click(); URL.revokeObjectURL(a.href);
}

/* ---------- backup, restore, reset ---------- */
async function restore(file) {
  try {
    const d = JSON.parse(await file.text());
    if (!Array.isArray(d.peserta) || !Array.isArray(d.hadiah) || !Array.isArray(d.winners)) throw 0;
    const r = await Swal.fire({ title: 'Restore data?', text: 'Data saat ini akan diganti dengan isi backup.', icon: 'warning', showCancelButton: true, confirmButtonText: 'Restore' });
    if (!r.isConfirmed) return;
    S = { peserta: d.peserta, hadiah: d.hadiah, winners: d.winners }; save(true); render();
    Swal.fire('Berhasil', 'Data berhasil direstore.', 'success');
  } catch (e) { Swal.fire('File tidak valid', 'File backup JSON tidak dikenali.', 'error'); }
}
async function reset() {
  const r = await Swal.fire({
    title: 'RESET DATA', icon: 'warning', showCancelButton: true, confirmButtonColor: '#d24a4a', confirmButtonText: 'Hapus semua data',
    html: 'Seluruh master peserta, master hadiah, dan riwayat undian akan <b>dihapus permanen</b>.<br>Ketik <b>RESET</b> untuk melanjutkan.',
    input: 'text', inputValidator: v => v !== 'RESET' ? 'Ketik RESET dengan huruf besar' : null
  });
  if (!r.isConfirmed) return;
  clearInterval(tm); tm = null; localStorage.removeItem(KEY); S = { peserta: [], hadiah: [], winners: [] }; render();
  Swal.fire('Selesai', 'Seluruh data telah dihapus.', 'success');
}

/* ---------- events ---------- */
document.querySelectorAll('nav [data-p]').forEach(b => b.onclick = () => {
  document.querySelectorAll('nav button').forEach(x => x.classList.remove('on')); b.classList.add('on');
  document.querySelectorAll('.page').forEach(p => p.classList.toggle('on', p.id === b.dataset.p)); render();
});
$('#fPeserta').onchange = e => { if (e.target.files[0]) doImport('peserta', e.target.files[0]); e.target.value = ''; };
$('#fHadiah').onchange = e => { if (e.target.files[0]) doImport('hadiah', e.target.files[0]); e.target.value = ''; };
$('#fRestore').onchange = e => { if (e.target.files[0]) restore(e.target.files[0]); e.target.value = ''; };
$('#selH').onchange = renderDraw;
$('#btnStart').onclick = start; $('#btnStop').onclick = stop;
$('#btnFs').onclick = () => document.fullscreenElement ? document.exitFullscreen() : $('#draw').requestFullscreen();
['#fQ', '#fH', '#fS'].forEach(s => $(s).oninput = renderWin);
$('#xXlsx').onclick = () => exportWin('xlsx'); $('#xCsv').onclick = () => exportWin('csv'); $('#xPdf').onclick = () => exportWin('pdf');
$('#slots').onclick = e => { const b = e.target.closest('button[data-n]'); if (b) mark(b.dataset.n, b.dataset.s); };
$('#btnBackup').onclick = () => { download('backup-undian-' + new Date().toISOString().slice(0, 10) + '.json', JSON.stringify(S, null, 2), 'application/json'); toast('Backup berhasil diunduh'); };
$('#btnReset').onclick = reset;
render();
