const CFG = {
  PENGATURAN: 'Pengaturan',
  REKAP: 'Rekap',
  HEADER_ROW: 8,      // baris judul kolom tabel transaksi
  FIRST_ROW: 9,       // baris saldo bulan lalu (otomatis)
  INPUT_ROW: 10,      // baris pertama untuk input manual
  LAST_ROW: 300,      // batas baris transaksi
  MAX_KATEGORI: 20,   // maksimal kategori per jenis di sheet Pengaturan
  BULAN: [
    { label: 'November',  singkat: 'Nov', m: 10, y: 2026 },
    { label: 'Desember',  singkat: 'Des', m: 11, y: 2026 },
    { label: 'Januari',   singkat: 'Jan', m: 0,  y: 2027 },
    { label: 'Februari',  singkat: 'Feb', m: 1,  y: 2027 },
    { label: 'Maret',     singkat: 'Mar', m: 2,  y: 2027 }
  ],
  KAT_PEMASUKAN: [
    'Kiriman Orang Tua', 'Beasiswa', 'Gaji Part-time', 'Hasil Usaha/Freelance',
    'Uang Saku Lainnya', 'Hadiah/Lain-lain'
  ],
  KAT_PENGELUARAN: [
    'Makan & Minum', 'Transportasi', 'Kos/Sewa', 'Listrik & Air', 'Pulsa & Internet',
    'UKT/SPP', 'Buku & Fotokopi', 'Tugas/Praktikum', 'Kesehatan', 'Hiburan',
    'Belanja', 'Organisasi', 'Tabungan', 'Lain-lain'
  ]
};

const RUPIAH = '"Rp" #,##0';

function namaSheet(b) { return b.singkat + ' ' + b.y; }
function adalahSheetBulan(nama) { return CFG.BULAN.some(b => namaSheet(b) === nama); }

/* ===================== MENU ===================== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 Keuangan')
    .addItem('Buat semua sheet (Nov–Mar)', 'setupKeuangan')
    .addItem('Segarkan dropdown kategori', 'segarkanDropdown')
    .addToUi();
}

===================== SETUP ===================== 

function setupKeuangan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  buatPengaturan(ss);
  CFG.BULAN.forEach((b, i) => buatSheetBulan(ss, b, i));
  buatRekap(ss);

  // Urutkan: Rekap, Pengaturan, lalu bulan-bulan
  const urutan = [CFG.REKAP, CFG.PENGATURAN].concat(CFG.BULAN.map(namaSheet));
  urutan.forEach((n, i) => {
    const s = ss.getSheetByName(n);
    if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(i + 1); }
  });

  // Hapus sheet kosong bawaan
  ss.getSheets().forEach(s => {
    if (['Sheet1', 'Sheet 1', 'Lembar1', 'Lembar 1'].indexOf(s.getName()) >= 0 &&
        s.getLastRow() === 0 && ss.getSheets().length > 1) {
      ss.deleteSheet(s);
    }
  });

  ss.setActiveSheet(ss.getSheetByName(CFG.REKAP));
  SpreadsheetApp.getUi().alert('Selesai! Sheet keuangan Nov–Mar sudah dibuat.\n\n' +
    'Isi saldo awal (jika ada) dan atur kategori di sheet "Pengaturan".');
}

===================== SHEET PENGATURAN =====================

function buatPengaturan(ss) {
  if (ss.getSheetByName(CFG.PENGATURAN)) return;
  const sh = ss.insertSheet(CFG.PENGATURAN);

  sh.getRange('A1:B1').setValues([['Kategori Pemasukan', 'Kategori Pengeluaran']])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.getRange(2, 1, CFG.KAT_PEMASUKAN.length, 1).setValues(CFG.KAT_PEMASUKAN.map(x => [x]));
  sh.getRange(2, 2, CFG.KAT_PENGELUARAN.length, 1).setValues(CFG.KAT_PENGELUARAN.map(x => [x]));

  sh.getRange('D1').setValue('Saldo Awal (sebelum November)')
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
  sh.getRange('D2').setValue(0).setNumberFormat(RUPIAH).setBackground('#fff8e1');
  sh.getRange('D4').setValue('Petunjuk:').setFontWeight('bold');
  sh.getRange('D5').setValue('• Ubah/tambah kategori di kolom A dan B (maks. ' + CFG.MAX_KATEGORI + ' per jenis).');
  sh.getRange('D6').setValue('• Isi saldo awal di D2 bila sudah punya uang sebelum November.');
  sh.getRange('D7').setValue('• Perubahan kategori langsung ikut ke dropdown dan rekap tiap bulan.');

  sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 200); sh.setColumnWidth(3, 30); sh.setColumnWidth(4, 230);
  sh.setTabColor('#9e9e9e');
}

===================== SHEET BULANAN =====================

function buatSheetBulan(ss, b, idx) {
  const nama = namaSheet(b);
  if (ss.getSheetByName(nama)) return;
  const sh = ss.insertSheet(nama);
  const R = CFG;

  // Judul
  sh.getRange('A1:F1').merge()
    .setValue('CATATAN KEUANGAN – ' + b.label.toUpperCase() + ' ' + b.y)
    .setFontSize(16).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // Ringkasan
  sh.getRange('A3:A6').setValues([['Saldo Awal'], ['Pemasukan Bulan Ini'], ['Total Pengeluaran'], ['Saldo Akhir']])
    .setFontWeight('bold');
  sh.getRange('B3:B6').setFormulas([
    ['=E' + R.FIRST_ROW],
    ['=SUMIF($B$' + R.FIRST_ROW + ':$B$' + R.LAST_ROW + ',"Pemasukan",$E$' + R.FIRST_ROW + ':$E$' + R.LAST_ROW + ')-B3'],
    ['=SUMIF($B$' + R.FIRST_ROW + ':$B$' + R.LAST_ROW + ',"Pengeluaran",$E$' + R.FIRST_ROW + ':$E$' + R.LAST_ROW + ')'],
    ['=B3+B4-B5']
  ]).setNumberFormat(RUPIAH);
  sh.getRange('A3:B6').setBackground('#e8f0fe').setBorder(true, true, true, true, true, true);
  sh.getRange('A6:B6').setFontWeight('bold').setBackground('#d2e3fc');

  sh.getRange('D3:D5').setValues([['Status'], ['Rasio Pengeluaran'], ['Jumlah Transaksi']]).setFontWeight('bold');
  sh.getRange('E3:E5').setFormulas([
    ['=IF(B6<0,"⚠ DEFISIT","✔ Aman")'],
    ['=IF(B3+B4=0,0,B5/(B3+B4))'],
    ['=COUNTA($C$' + R.INPUT_ROW + ':$C$' + R.LAST_ROW + ')']
  ]);
  sh.getRange('E4').setNumberFormat('0%');
  sh.getRange('D3:E5').setBackground('#e8f0fe').setBorder(true, true, true, true, true, true);
  sh.getRange('E3:E5').setHorizontalAlignment('right');

  // Header tabel transaksi
  sh.getRange(R.HEADER_ROW, 1, 1, 6)
    .setValues([['Tanggal', 'Jenis', 'Kategori', 'Keterangan', 'Jumlah (Rp)', 'Saldo Berjalan']])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff').setHorizontalAlignment('center');

  // Baris saldo bulan lalu (otomatis)
  const sumber = idx === 0
    ? "='" + R.PENGATURAN + "'!$D$2"
    : "='" + namaSheet(R.BULAN[idx - 1]) + "'!$B$6";
  sh.getRange(R.FIRST_ROW, 1, 1, 5).setValues([[
    new Date(b.y, b.m, 1), 'Pemasukan', 'Saldo Bulan Lalu',
    idx === 0 ? 'Saldo awal (dari sheet Pengaturan)' : 'Sisa uang bulan sebelumnya (otomatis)', ''
  ]]);
  sh.getRange(R.FIRST_ROW, 5).setFormula(sumber);
  sh.getRange(R.FIRST_ROW, 1, 1, 6).setBackground('#eeeeee').setFontStyle('italic');
  sh.getRange(R.FIRST_ROW, 1, 1, 6).protect().setWarningOnly(true);

  // Saldo berjalan
  sh.getRange(R.FIRST_ROW, 6, R.LAST_ROW - R.FIRST_ROW + 1, 1).setFormula(
    '=IF(E' + R.FIRST_ROW + '="","",SUMIF($B$' + R.FIRST_ROW + ':B' + R.FIRST_ROW + ',"Pemasukan",$E$' + R.FIRST_ROW + ':E' + R.FIRST_ROW +
    ')-SUMIF($B$' + R.FIRST_ROW + ':B' + R.FIRST_ROW + ',"Pengeluaran",$E$' + R.FIRST_ROW + ':E' + R.FIRST_ROW + '))'
  );

  // Format kolom
  const nRows = R.LAST_ROW - R.FIRST_ROW + 1;
  sh.getRange(R.FIRST_ROW, 1, nRows, 1).setNumberFormat('dd/mm/yyyy').setHorizontalAlignment('center');
  sh.getRange(R.FIRST_ROW, 5, nRows, 2).setNumberFormat(RUPIAH);
  sh.getRange(R.HEADER_ROW, 1, nRows + 1, 6).setBorder(true, true, true, true, true, true, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);

  // Validasi input
  const awal = new Date(b.y, b.m, 1), akhir = new Date(b.y, b.m + 1, 0);
  sh.getRange(R.INPUT_ROW, 1, R.LAST_ROW - R.INPUT_ROW + 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireDateBetween(awal, akhir)
      .setAllowInvalid(false).setHelpText('Isi tanggal di bulan ' + b.label + ' ' + b.y).build());
  sh.getRange(R.INPUT_ROW, 2, R.LAST_ROW - R.INPUT_ROW + 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Pemasukan', 'Pengeluaran'], true)
      .setAllowInvalid(false).build());
  sh.getRange(R.INPUT_ROW, 5, R.LAST_ROW - R.INPUT_ROW + 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0)
      .setAllowInvalid(false).setHelpText('Isi nominal positif (tanpa tanda minus)').build());

  // Rekap per kategori (kanan)
  buatRekapKategori(sh);

  // Warna baris berdasarkan jenis + saldo negatif merah
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B10="Pemasukan"')
      .setBackground('#e6f4ea').setRanges([sh.getRange('A10:F' + R.LAST_ROW)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$B10="Pengeluaran"')
      .setBackground('#fce8e6').setRanges([sh.getRange('A10:F' + R.LAST_ROW)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0)
      .setFontColor('#d93025').setRanges([sh.getRange('B6'), sh.getRange('F9:F' + R.LAST_ROW)]).build()
  ]);

  // Lebar kolom & tampilan
  [100, 110, 170, 240, 130, 130, 20, 20, 190, 120, 20, 190, 120]
    .forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(R.HEADER_ROW);
  sh.setTabColor(['#1a73e8', '#34a853', '#fbbc04', '#ea4335', '#9334e6'][idx % 5]);
}

function buatRekapKategori(sh) {
  const R = CFG, n = R.MAX_KATEGORI;
  const rng = '$E$' + R.INPUT_ROW + ':$E$' + R.LAST_ROW;
  const jns = '$B$' + R.INPUT_ROW + ':$B$' + R.LAST_ROW;
  const kat = '$C$' + R.INPUT_ROW + ':$C$' + R.LAST_ROW;

  sh.getRange('I7').setValue('REKAP PER KATEGORI').setFontWeight('bold');
  sh.getRange('I8:J8').setValues([['Kategori Pemasukan', 'Total']]);
  sh.getRange('L8:M8').setValues([['Kategori Pengeluaran', 'Total']]);
  sh.getRange('I8:J8').setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
  sh.getRange('L8:M8').setFontWeight('bold').setBackground('#ea4335').setFontColor('#ffffff');

  const fIn = [], fOut = [];
  for (let k = 0; k < n; k++) {
    const r = R.FIRST_ROW + k;       // baris di sheet bulan
    const p = 2 + k;                 // baris di sheet Pengaturan
    fIn.push([
      '=IF(' + R.PENGATURAN + '!A' + p + '="","",' + R.PENGATURAN + '!A' + p + ')',
      '=IF(I' + r + '="","",SUMIFS(' + rng + ',' + jns + ',"Pemasukan",' + kat + ',I' + r + '))'
    ]);
    fOut.push([
      '=IF(' + R.PENGATURAN + '!B' + p + '="","",' + R.PENGATURAN + '!B' + p + ')',
      '=IF(L' + r + '="","",SUMIFS(' + rng + ',' + jns + ',"Pengeluaran",' + kat + ',L' + r + '))'
    ]);
  }
  sh.getRange(R.FIRST_ROW, 9, n, 2).setFormulas(fIn);
  sh.getRange(R.FIRST_ROW, 12, n, 2).setFormulas(fOut);
  sh.getRange(R.FIRST_ROW, 10, n, 1).setNumberFormat(RUPIAH);
  sh.getRange(R.FIRST_ROW, 13, n, 1).setNumberFormat(RUPIAH);
}

===================== SHEET REKAP =====================

function buatRekap(ss) {
  if (ss.getSheetByName(CFG.REKAP)) return;
  const sh = ss.insertSheet(CFG.REKAP);

  sh.getRange('A1:F1').merge().setValue('REKAP KEUANGAN NOVEMBER – MARET')
    .setFontSize(16).setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  sh.getRange('A3:F3').setValues([['Bulan', 'Saldo Awal', 'Pemasukan', 'Pengeluaran', 'Saldo Akhir', 'Status']])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff').setHorizontalAlignment('center');

  const rows = CFG.BULAN.map(b => {
    const q = "'" + namaSheet(b) + "'!";
    return [b.label + ' ' + b.y, '=' + q + 'B3', '=' + q + 'B4', '=' + q + 'B5', '=' + q + 'B6', '=' + q + 'E3'];
  });
  sh.getRange(4, 1, rows.length, 6).setFormulas(rows);

  const last = 3 + rows.length, tot = last + 1;
  sh.getRange(tot, 1, 1, 6).setFormulas([[
    'TOTAL', '=B4', '=SUM(C4:C' + last + ')', '=SUM(D4:D' + last + ')', '=E' + last, ''
  ]]).setFontWeight('bold').setBackground('#d2e3fc');
  sh.getRange(4, 2, rows.length + 1, 4).setNumberFormat(RUPIAH);
  sh.getRange(3, 1, rows.length + 2, 6)
    .setBorder(true, true, true, true, true, true, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);

  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0)
      .setFontColor('#d93025').setRanges([sh.getRange('E4:E' + tot)]).build()
  ]);

  // Grafik pemasukan vs pengeluaran per bulan
  const chart = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange('A3:A' + last))
    .addRange(sh.getRange('C3:D' + last))
    .setNumHeaders(1)
    .setOption('title', 'Pemasukan vs Pengeluaran per Bulan')
    .setOption('colors', ['#34a853', '#ea4335'])
    .setPosition(tot + 2, 1, 0, 0)
    .setOption('width', 640).setOption('height', 320)
    .build();
  sh.insertChart(chart);

  [150, 130, 130, 130, 130, 110].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setTabColor('#34a853');
}

===================== DROPDOWN KATEGORI =====================

Dipanggil otomatis saat kolom "Jenis" diubah: dropdown Kategori menyesuaikan.

function onEdit(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  if (!adalahSheetBulan(sh.getName())) return;
  const r = e.range;
  if (r.getColumn() > 2 || r.getLastColumn() < 2) return;   // hanya jika kolom B ikut berubah

  const start = Math.max(r.getRow(), CFG.INPUT_ROW);
  const end = Math.min(r.getLastRow(), CFG.LAST_ROW);
  for (let row = start; row <= end; row++) terapkanKategori(sh, row);
}

function terapkanKategori(sh, row) {
  const ss = sh.getParent();
  const set = ss.getSheetByName(CFG.PENGATURAN);
  if (!set) return;
  const jenis = sh.getRange(row, 2).getValue();
  const sel = sh.getRange(row, 3);
  let sumber = null;

  if (jenis === 'Pemasukan') sumber = set.getRange(2, 1, CFG.MAX_KATEGORI, 1);
  else if (jenis === 'Pengeluaran') sumber = set.getRange(2, 2, CFG.MAX_KATEGORI, 1);

  if (!sumber) { sel.clearDataValidations(); return; }

  sel.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(sumber, true).setAllowInvalid(false).build());

  // Kosongkan kategori jika tidak cocok dengan jenis yang dipilih
  const val = sel.getValue();
  if (val !== '') {
    const daftar = sumber.getValues().map(x => x[0]);
    if (daftar.indexOf(val) < 0) sel.clearContent();
  }
}

Terapkan ulang dropdown kategori ke semua baris yang sudah terisi (lewat menu). 

function segarkanDropdown() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  CFG.BULAN.forEach(b => {
    const sh = ss.getSheetByName(namaSheet(b));
    if (!sh) return;
    const jenis = sh.getRange(CFG.INPUT_ROW, 2, CFG.LAST_ROW - CFG.INPUT_ROW + 1, 1).getValues();
    jenis.forEach((v, i) => { if (v[0] !== '') terapkanKategori(sh, CFG.INPUT_ROW + i); });
  });
  SpreadsheetApp.getUi().alert('Dropdown kategori sudah diperbarui.');
}
