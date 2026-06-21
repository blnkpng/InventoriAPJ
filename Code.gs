/**
 * APJ INVENTORI - Code.gs Revisi v14.8 Print Input Stok 58mm Multi Barang
 * Fokus revisi v3:
 * 1) Kategori Produksi tidak lagi menyimpan SEMUA bahan ke TRANSAKSI_OUTPUT_PRODUKSI.
 * 2) Produksi/Lauk jadi barang jadi untuk dijual: masuk 1 baris per produksi ke TRANSAKSI_OUTPUT_PRODUKSI.
 * 3) Bahan mentah/bahan baku yang dipakai produksi otomatis masuk sebagai pengeluaran ke TRANSAKSI_OUTPUT_STOK.
 * 4) ID Produksi diambil dari MASTER_BARANG_PRODUKSI Kolom A.
 * 5) ID Bahan pengeluaran dicari dari STOK_AKHIR berdasarkan nama Bahan Baku.
 * 6) Tambahan endpoint getStokProduksiReport untuk halaman Lihat Stok kategori Produksi.
 * 7) Revisi v5: MASTER_BAHAN_MENTAH mendukung Parent-Child preparasi ayam.
 *    - Input: Parent tampil, Child disembunyikan dan muncul otomatis sebagai hasil preparasi.
 *    - Output/Opname/Lihat Stok: Parent disembunyikan, Child tampil sebagai stok operasional.
 */

var CONFIG = {
  SHEET_ID: '1hsmlb_kviAbsn7-HEZDnjz8FPxwh2eJBxrOmnfkvTCg',
  TZ: 'Asia/Jakarta',
  SESSION_TTL_SECONDS: 21600,
  DEFAULT_LOCATION: 'PUSAT',
  SHEETS: {
    USER: 'USER',
    PERMISSION: 'LEVEL_PERMISSION',
    STOK_AKHIR: 'STOK_AKHIR',
    OUTLET: 'OUTLET',
    MASTER_PRODUKSI: 'MASTER_BARANG_PRODUKSI',
    MASTER_PRODUK_OUTLET: 'MASTER_PRODUK_OUTLET',
    INPUT_STOK: 'TRANSAKSI_INPUT_STOK',
    OUTPUT_STOK: 'TRANSAKSI_OUTPUT_STOK',
    OUTPUT_PRODUKSI: 'TRANSAKSI_OUTPUT_PRODUKSI',
    TRANSFER_PRODUKSI: 'TRANSFER_PRODUKSI',
    STOK_OPNAME: 'STOK_OPNAME',
    PRODUK_OUTLET_KELUAR: 'PRODUK_OUTLET_KELUAR',
    OPNAME_PRODUK_OUTLET: 'OPNAME_PRODUK_OUTLET',
    AUDIT_LOG: 'AUDIT_LOG'
  }
};

var HEADER_INPUT_STOK = [
  'ID Transaksi', 'Tanggal', 'Petugas', 'Kategori', 'ID Barang', 'Nama Barang',
  'Qty Beli', 'Satuan Beli', 'Qty Stok', 'Satuan Stok'
];

var HEADER_OUTPUT_STOK = [
  'ID Transaksi', 'Tanggal', 'Petugas', 'Kategori', 'ID Barang', 'Nama Barang',
  'Qty Keluar', 'Satuan', 'Tujuan', 'Outlet', 'PIC'
];

var HEADER_OUTPUT_PRODUKSI = [
  'ID Transaksi', 'Tanggal', 'Petugas', 'ID Barang Produksi', 'Produksi',
  'Qty Produksi', 'Satuan Produksi'
];

var HEADER_TRANSFER_PRODUKSI = [
  'ID Transfer', 'Tanggal', 'Petugas', 'Outlet', 'PIC Penerima',
  'ID Produk', 'Nama Produk', 'Qty Transfer', 'Satuan', 'Catatan'
];

var HEADER_PRODUK_OUTLET_KELUAR = [
  'ID Transaksi', 'Tanggal', 'Outlet', 'Petugas', 'ID Produk', 'Nama Produk',
  'Qty Keluar', 'Satuan', 'Keterangan', 'Timestamp'
];

var HEADER_OPNAME_PRODUK_OUTLET = [
  'ID Opname', 'Tanggal', 'Outlet', 'Petugas', 'ID Produk', 'Nama Produk',
  'Stok Sistem', 'Stok Fisik', 'Selisih', 'Satuan', 'Alasan', 'Timestamp'
];

var HEADER_AUDIT_LOG = [
  'Timestamp', 'Action', 'Status', 'User', 'Level', 'Outlet', 'Message', 'Request ID'
];

function doPost(e) {
  try {
    var requestData = JSON.parse(e.postData.contents || '{}');
    var action = requestData.action;
    var ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

    if (action === 'login') return handleLogin_(ss, requestData);

    requestData._user = getSessionProfile_(ss, requestData) || getLegacyProfileFromRequest_(ss, requestData);

    // General: menarik master barang biasa + stok akhir.
    if (action === 'getBarang') return handleGetBarang_(ss, requestData);

    // Meta form Output Stok untuk dropdown Outlet dan PIC.
    if (action === 'getOutputInit') return handleGetOutputInit_(ss);

    // Data produk siap transfer + outlet/PIC.
    if (action === 'getTransferProduksiInit') return handleGetTransferProduksiInit_(ss, requestData);

    // v14.3: ambil data transfer yang sudah tersimpan untuk cetak Surat Jalan A4/58mm.
    if (action === 'getTransferPrintData') return handleGetTransferPrintData_(ss, requestData);

    // v14.7: opsi dan data cetak label input stok 58mm.
    if (action === 'getInputPrintOptions') return handleGetInputPrintOptions_(ss, requestData);
    if (action === 'getInputPrintData') return handleGetInputPrintData_(ss, requestData);

    // Data produksi dikelompokkan per nama produksi/lauk.
    if (action === 'getBarangProduksi') return handleGetBarangProduksi_(ss, requestData);

    // BARU v3: khusus halaman Lihat Stok untuk menampilkan stok lauk/produksi secara unik.
    if (action === 'getStokProduksiReport') return handleGetStokProduksiReport_(ss, requestData);

    // v8B: laporan langsung dari STOK_AKHIR header baru (Lokasi-aware).
    if (action === 'getStokAkhirReport') return handleGetStokAkhirReport_(ss, requestData);

    // DASHBOARD v4: data dashboard berbeda antara Owner dan role operasional.
    if (action === 'getDashboardData') return handleGetDashboardData_(ss, requestData);

    // ADMIN: hanya session Owner/SuperAdmin, tidak menerima fallback legacy.
    if (action === 'getAdminData') return handleGetAdminData_(ss, requestData);
    if (action === 'adminUpdatePassword') return handleAdminUpdatePassword_(ss, requestData);
    if (action === 'adminSavePermissions') return handleAdminSavePermissions_(ss, requestData);

    // RIWAYAT v13: gabungan transaksi input, output, produksi, transfer, dan opname.
    if (action === 'getRiwayatTransaksi') return handleGetRiwayatTransaksi_(ss, requestData);

    // PRODUK OUTLET v14: stok outlet, produk terjual, dan opname outlet.
    if (action === 'getProdukOutletData') return handleGetProdukOutletData_(ss, requestData);
    if (action === 'simpanProdukOutletKeluar') return handleSimpanProdukOutletKeluar_(ss, requestData);
    if (action === 'simpanProdukOutletOpname') return handleSimpanProdukOutletOpname_(ss, requestData);

    // REVISI v5: simpan input stok biasa + input preparasi ayam dalam 1 request.
    if (action === 'simpanInputStokPreparasi') return handleSimpanInputStokPreparasi_(ss, requestData);

    // General: simpan transaksi input/output/opname lama.
    if (action === 'simpanTransaksi') return handleSimpanTransaksi_(ss, requestData);

    // REVISI v2: simpan produksi jadi + pemakaian bahan dalam 1 ID transaksi.
    if (action === 'simpanOutputProduksi') return handleSimpanOutputProduksiV2_(ss, requestData);

    // REVISI v10: simpan transfer produk/lauk dari PUSAT ke outlet.
    if (action === 'simpanTransferProduksi') return handleSimpanTransferProduksi_(ss, requestData);

    if (action === 'getUsers') return dataResponse_({ success: true, data: getUserNames_(ss) });

    return dataResponse_({ success: false, message: 'Aksi tidak valid.' });
  } catch (error) {
    return dataResponse_({ success: false, message: 'Error Server: ' + error.message });
  }
}


/**
 * REVISI v7 - Permission dinamis dari header LEVEL_PERMISSION.
 * Tujuan: setelah ini cukup edit sheet LEVEL_PERMISSION, tidak perlu ubah Code.gs.
 * Header yang dikenali fleksibel, misalnya:
 * - BahanMentah / Bahan Mentah
 * - BahanBaku / Bahan Baku
 * - Lihat Stok / LihatStok
 */
function defaultPermissions_(fillValue) {
  var v = fillValue || 'N';
  return {
    bahanMentah: v,
    bahanBaku: v,
    kemasan: v,
    minuman: v,
    kebersihan: v,
    opname: v,
    produksi: v,
    transferProduksi: v,
    lihatStok: v,
    produkOutlet: v
  };
}

function readPermissionsForLevel_(ss, level) {
  var isOwner = isOwnerRole_(level);
  var fallback = defaultPermissions_(isOwner ? 'Y' : 'N');
  var permSheet = ss.getSheetByName(CONFIG.SHEETS.PERMISSION);
  if (!permSheet || permSheet.getLastRow() < 2) return fallback;

  var data = permSheet.getDataRange().getValues();
  var headers = data[0] || [];
  var colToKey = {};
  for (var c = 0; c < headers.length; c++) {
    var key = permissionKeyFromHeader_(headers[c]);
    if (key) colToKey[c] = key;
  }

  for (var r = 1; r < data.length; r++) {
    if (normalizeKey_(data[r][0]) !== normalizeKey_(level)) continue;

    var permissions = defaultPermissions_('N');
    for (var col in colToKey) {
      permissions[colToKey[col]] = asYN_(data[r][Number(col)]);
    }

    // Owner/SuperAdmin tetap full access, kecuali nanti sengaja dibuat role non-owner.
    if (isOwner) {
      Object.keys(permissions).forEach(function(k) { permissions[k] = 'Y'; });
    }
    return permissions;
  }

  return fallback;
}

function permissionKeyFromHeader_(header) {
  var h = normalizeHeader_(header);
  if (h === 'bahanmentah' || h === 'mentah') return 'bahanMentah';
  if (h === 'bahanbaku' || h === 'baku') return 'bahanBaku';
  if (h === 'kemasan') return 'kemasan';
  if (h === 'minuman') return 'minuman';
  if (h === 'kebersihan' || h === 'bersih') return 'kebersihan';
  if (h === 'opname' || h === 'stokopname') return 'opname';
  if (h === 'produksi') return 'produksi';
  if (h === 'transferproduksi' || h === 'transfer' || h === 'kirimproduksi') return 'transferProduksi';
  if (h === 'lihatstok' || h === 'lihat') return 'lihatStok';
  if (h === 'produkoutlet' || h === 'productoutlet' || h === 'stokprodukoutlet' || h === 'outletproduk') return 'produkOutlet';
  return '';
}

function buildHeaderMap_(headerRow) {
  var map = {};
  for (var i = 0; i < headerRow.length; i++) {
    var key = normalizeHeader_(headerRow[i]);
    if (key && typeof map[key] === 'undefined') map[key] = i;
  }
  return map;
}

function findHeaderIndex_(headerMap, aliases, fallback) {
  for (var i = 0; i < aliases.length; i++) {
    var key = normalizeHeader_(aliases[i]);
    if (typeof headerMap[key] !== 'undefined') return headerMap[key];
  }
  return fallback;
}

function normalizeHeader_(value) {
  return asText_(value).toLowerCase().replace(/[^a-z0-9]/g, '');
}

function masterSheetPermissionKey_(sheetName) {
  var key = normalizeHeader_(sheetName);
  if (key === 'masterbahanmentah') return 'bahanMentah';
  if (key === 'masterbahanbaku') return 'bahanBaku';
  if (key === 'masterkemasan') return 'kemasan';
  if (key === 'masterminuman') return 'minuman';
  if (key === 'masterkebersihan') return 'kebersihan';
  return '';
}

function isAllowedMasterSheet_(sheetName) {
  return !!masterSheetPermissionKey_(sheetName);
}

function profileHasPermission_(profile, permissionKey) {
  if (!profile) return false;
  if (isOwnerRole_(profile.level)) return true;
  return permissionYes_(profile.permissions || {}, permissionKey);
}

function profileCanReadCategory_(profile, permissionKey) {
  return profileHasPermission_(profile, permissionKey) || profileHasPermission_(profile, 'lihatStok');
}

function profileHasAnyStockCategory_(profile) {
  var keys = ['bahanMentah', 'bahanBaku', 'kemasan', 'minuman', 'kebersihan'];
  for (var i = 0; i < keys.length; i++) {
    if (profileHasPermission_(profile, keys[i])) return true;
  }
  return false;
}

function getSessionCacheKey_(token) {
  return 'APJ_SESSION_' + asText_(token);
}

function getSessionTokenFromRequest_(requestData) {
  return asText_(requestData.sessionToken || requestData.token || requestData.authToken);
}

function createSessionForUser_(user) {
  var token = Utilities.getUuid();
  var payload = {
    username: user.username,
    nama: user.nama,
    createdAt: Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss')
  };
  CacheService.getScriptCache().put(
    getSessionCacheKey_(token),
    JSON.stringify(payload),
    CONFIG.SESSION_TTL_SECONDS
  );
  return token;
}

function getSessionProfile_(ss, requestData) {
  var token = getSessionTokenFromRequest_(requestData || {});
  if (!token) return null;

  var raw = CacheService.getScriptCache().get(getSessionCacheKey_(token));
  if (!raw) return null;

  var session = {};
  try {
    session = JSON.parse(raw);
  } catch (err) {
    return null;
  }

  var profile = getUserProfileByUsername_(ss, session.username);
  if (!profile && session.nama) profile = getUserProfileByName_(ss, session.nama);
  if (!profile) return null;

  profile.permissions = readPermissionsForLevel_(ss, profile.level);
  return profile;
}

function getLegacyProfileFromRequest_(ss, requestData) {
  requestData = requestData || {};
  var name = asText_(requestData.userName || requestData.petugas || requestData.nama || '');
  var username = asText_(requestData.username || '');

  if (!name && requestData.data && requestData.data.length && requestData.data[0]) {
    name = asText_(requestData.data[0][1]);
  }
  if (!name && requestData.inputRows && requestData.inputRows.length && requestData.inputRows[0]) {
    name = asText_(requestData.inputRows[0][1]);
  }
  if (!name && requestData.outputRows && requestData.outputRows.length && requestData.outputRows[0]) {
    name = asText_(requestData.outputRows[0][1]);
  }
  if (!name && requestData.produksiRows && requestData.produksiRows.length && requestData.produksiRows[0]) {
    name = asText_(requestData.produksiRows[0].petugas);
  }
  if (!name && requestData.bahanRows && requestData.bahanRows.length && requestData.bahanRows[0]) {
    name = asText_(requestData.bahanRows[0].petugas);
  }

  var profile = username ? getUserProfileByUsername_(ss, username) : null;
  if (!profile && name) profile = getUserProfileByName_(ss, name);
  if (!profile) return null;

  profile.permissions = readPermissionsForLevel_(ss, profile.level);
  return profile;
}

function getUserProfileByUsername_(ss, username) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.USER);
  if (!sheet || sheet.getLastRow() < 2) return null;

  var data = sheet.getDataRange().getValues();
  var hm = buildHeaderMap_(data[0]);
  var ixUsername = findHeaderIndex_(hm, ['username', 'user', 'namauser'], 1);
  var ixNama = findHeaderIndex_(hm, ['nama', 'namakaryawan', 'petugas'], 3);
  var ixLevel = findHeaderIndex_(hm, ['level', 'role', 'akses'], 4);
  var ixOutlet = findHeaderIndex_(hm, ['outlet', 'lokasi', 'cabang'], 5);
  var userKey = normalizeKey_(username);

  for (var i = 1; i < data.length; i++) {
    if (normalizeKey_(data[i][ixUsername]) !== userKey) continue;
    var level = asText_(data[i][ixLevel]);
    var outlet = asText_(data[i][ixOutlet]);
    if (!outlet && isOwnerRole_(level)) outlet = 'ALL';
    return {
      username: asText_(data[i][ixUsername]),
      nama: asText_(data[i][ixNama]),
      level: level,
      outlet: outlet || ''
    };
  }
  return null;
}

/**
 * LOGIN v6.
 * Revisi penting:
 * 1) Username dibuat case-insensitive agar "anandaarifp" tetap cocok dengan "Anandaarifp" di sheet.
 * 2) Password dibandingkan sebagai teks murni, tanpa aturan minimal 8 karakter dari sistem.
 *    Jadi password pendek tetap bisa login selama nilainya sama dengan USER kolom C.
 * 3) Mapping LEVEL_PERMISSION diperbaiki sesuai sheet saat ini:
 *    G = Opname, H = Produksi, I = Lihat Stok.
 */
function handleLogin_(ss, requestData) {
  var usernameInput = asText_(requestData.username);
  var passwordInput = asText_(requestData.password);
  var usernameKey = normalizeKey_(usernameInput);

  var userSheet = ss.getSheetByName(CONFIG.SHEETS.USER);
  var permSheet = ss.getSheetByName(CONFIG.SHEETS.PERMISSION);

  if (!userSheet || !permSheet) {
    return dataResponse_({ success: false, message: 'Tabel USER / LEVEL_PERMISSION tidak ditemukan.' });
  }

  var users = userSheet.getDataRange().getValues();
  if (users.length < 2) return dataResponse_({ success: false, message: 'Tabel USER masih kosong.' });

  var headerMap = buildHeaderMap_(users[0]);
  var ixUsername = findHeaderIndex_(headerMap, ['username', 'user', 'namauser'], 1);
  var ixPassword = findHeaderIndex_(headerMap, ['password', 'pass', 'sandi'], 2);
  var ixNama = findHeaderIndex_(headerMap, ['nama', 'namakaryawan', 'petugas'], 3);
  var ixLevel = findHeaderIndex_(headerMap, ['level', 'role', 'akses'], 4);
  var ixOutlet = findHeaderIndex_(headerMap, ['outlet', 'lokasi', 'cabang'], 5);

  var loggedInUser = null;

  for (var i = 1; i < users.length; i++) {
    var sheetUsername = asText_(users[i][ixUsername]);
    var sheetPassword = asText_(users[i][ixPassword]);

    // Username tidak sensitif huruf besar/kecil. Password mendukung mode lama dan hash SHA-256.
    if (normalizeKey_(sheetUsername) === usernameKey && passwordMatches_(passwordInput, sheetPassword)) {
      var level = asText_(users[i][ixLevel]);
      var outlet = asText_(users[i][ixOutlet]);
      if (!outlet && isOwnerRole_(level)) outlet = 'ALL';
      loggedInUser = {
        username: sheetUsername,
        nama: asText_(users[i][ixNama]),
        level: level,
        outlet: outlet || ''
      };
      break;
    }
  }

  if (!loggedInUser) {
    appendAuditLog_(ss, { action: 'login', username: usernameInput }, 'FAILED', 'Username atau password salah.');
    return dataResponse_({ success: false, message: 'Username atau password salah.' });
  }

  // Permission dibaca dinamis dari header LEVEL_PERMISSION.
  var userPermissions = readPermissionsForLevel_(ss, loggedInUser.level);
  var sessionToken = createSessionForUser_(loggedInUser);
  appendAuditLog_(ss, {
    action: 'login',
    username: loggedInUser.username,
    _user: {
      username: loggedInUser.username,
      nama: loggedInUser.nama,
      level: loggedInUser.level,
      outlet: loggedInUser.outlet
    }
  }, 'SUCCESS', 'Login berhasil.');

  return dataResponse_({
    success: true,
    message: 'Otentikasi Berhasil.',
    sessionToken: sessionToken,
    expiresInSeconds: CONFIG.SESSION_TTL_SECONDS,
    user: {
      username: loggedInUser.username,
      nama: loggedInUser.nama,
      level: loggedInUser.level,
      outlet: loggedInUser.outlet,
      permissions: userPermissions
    }
  });
}

/**
 * GENERAL: Dipakai Bahan Mentah/Baku/Kemasan/Minuman/Kebersihan.
 *
 * REVISI v5 - Context-aware untuk MASTER_BAHAN_MENTAH:
 * - context = 'input'       : tampilkan item normal + PARENT. CHILD disembunyikan dari dropdown utama,
 *                             tapi dikirim sebagai children agar muncul otomatis di form preparasi.
 * - context = 'operasional' : tampilkan item normal + CHILD. PARENT disembunyikan untuk Output, Opname,
 *                             dan Lihat Stok agar Ayam Potong 8/11 tidak dipakai langsung.
 * - tanpa context           : kompatibel mode lama, tampilkan semua item.
 */
function handleGetBarang_(ss, requestData) {
  var targetSheetName = requestData.sheetName;
  var context = normalizeKey_(requestData.context || requestData.mode || '');
  var permissionKey = masterSheetPermissionKey_(targetSheetName);
  if (!isAllowedMasterSheet_(targetSheetName)) {
    return dataResponse_({ success: false, message: 'Sheet master tidak diizinkan untuk endpoint ini.' });
  }
  if (requestData._user && !profileCanReadCategory_(requestData._user, permissionKey)) {
    return dataResponse_({ success: false, message: 'Akses kategori ini ditolak untuk user login.' });
  }
  var sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) return dataResponse_({ success: false, message: "Tabel Master '" + targetSheetName + "' tidak ditemukan." });

  var stokMaps = buildStockMaps_(ss);
  var data = sheet.getDataRange().getValues();
  var rawItems = [];
  var childrenByParent = {};

  for (var k = 1; k < data.length; k++) {
    if (!data[k][0] || !data[k][1]) continue;

    var item = buildMasterItem_(data[k], stokMaps);
    rawItems.push(item);

    if (item.parentId) {
      var parentKey = normalizeKey_(item.parentId);
      if (!childrenByParent[parentKey]) childrenByParent[parentKey] = [];
      childrenByParent[parentKey].push(item);
    }
  }

  var listBarang = [];
  for (var i = 0; i < rawItems.length; i++) {
    var current = rawItems[i];

    // Input: sembunyikan child dari dropdown utama.
    if (context === 'input') {
      if (current.isChild || current.tampilInput === 'N') continue;
    }

    // Output/Opname/Lihat: sembunyikan parent. Child tetap tampil walaupun Tampil Input = N.
    if (context === 'operasional' || context === 'output' || context === 'opname' || context === 'lihat') {
      if (current.isParent) continue;
    }

    var out = clonePlainItem_(current);
    var childList = childrenByParent[normalizeKey_(current.id)] || [];
    out.children = childList.map(clonePlainItem_);
    out.hasChildren = out.children.length > 0;
    listBarang.push(out);
  }

  return dataResponse_({ success: true, data: listBarang });
}

/** Membaca 1 baris master barang, termasuk kolom F-I untuk Parent-Child. */
function buildMasterItem_(row, stokMaps) {
  var id = asText_(row[0]);
  var nama = asText_(row[1]);
  var stokItem = getStockItem_(stokMaps, CONFIG.DEFAULT_LOCATION, id, nama);

  var tipeItem = asText_(row[5]).toUpperCase();        // Kolom F: PARENT / CHILD / kosong
  var parentId = asText_(row[6]);                      // Kolom G: Parent ID
  var tampilInput = asText_(row[7]).toUpperCase();     // Kolom H: Y / N / kosong
  var grupProses = asText_(row[8]);                    // Kolom I: contoh AYAM

  return {
    id: id,
    nama: nama,
    satuanBeli: asText_(row[2]) || '-',
    satuanStok: asText_(row[3]) || (stokItem ? stokItem.satuan : '-') || '-',
    satuanProduksi: asText_(row[4]) || '-',
    tipeItem: tipeItem,
    parentId: parentId,
    tampilInput: tampilInput,
    grupProses: grupProses,
    isParent: tipeItem === 'PARENT',
    isChild: tipeItem === 'CHILD' || !!parentId,
    stokTersedia: stokItem ? stokItem.qty : 0,
    lastUpdate: stokItem ? stokItem.lastUpdate : '',
    status: stokItem ? stokItem.status : '',
    kategoriStok: stokItem ? stokItem.kategori : ''
  };
}

function clonePlainItem_(item) {
  return {
    id: item.id,
    nama: item.nama,
    satuanBeli: item.satuanBeli,
    satuanStok: item.satuanStok,
    satuanProduksi: item.satuanProduksi,
    tipeItem: item.tipeItem,
    parentId: item.parentId,
    tampilInput: item.tampilInput,
    grupProses: item.grupProses,
    isParent: item.isParent,
    isChild: item.isChild,
    stokTersedia: item.stokTersedia,
    lastUpdate: item.lastUpdate,
    status: item.status,
    kategoriStok: item.kategoriStok
  };
}

/** Outlet dari sheet OUTLET dan PIC dari USER kolom D. */
function handleGetOutputInit_(ss) {
  return dataResponse_({
    success: true,
    outlets: getOutletNames_(ss),
    pics: getUserNames_(ss)
  });
}


/**
 * REVISI v14.1: init halaman Transfer Produk Outlet.
 * Daftar produk tetap dari MASTER_PRODUK_OUTLET.
 * Stok PUSAT dibaca dari STOK_AKHIR untuk SEMUA kategori, sehingga barang siap jual seperti Minuman & Barang Dagang bisa langsung ditransfer tanpa proses produksi.
 */
function handleGetTransferProduksiInit_(ss, requestData) {
  var profile = requestData._user;
  if (profile && !profileHasPermission_(profile, 'transferProduksi') && !profileHasPermission_(profile, 'produksi')) {
    return dataResponse_({ success: false, message: 'Akses Transfer Produk Outlet ditolak untuk user login.' });
  }
  var outlet = toSheetText_(requestData.outlet || requestData.outletTujuan);
  var produk = getTransferProdukList_(ss, outlet);
  return dataResponse_({
    success: true,
    outlets: getOutletNames_(ss),
    pics: getUserNames_(ss),
    produk: produk,
    outlet: outlet,
    defaultLocation: CONFIG.DEFAULT_LOCATION,
    menggunakanMasterOutlet: !!ss.getSheetByName(CONFIG.SHEETS.MASTER_PRODUK_OUTLET)
  });
}

/**
 * v14.1: daftar produk transfer dibatasi per outlet dari MASTER_PRODUK_OUTLET.
 * Jika outlet dipilih dan MASTER_PRODUK_OUTLET ada, maka yang tampil HANYA produk aktif outlet tersebut.
 * Stok dasar dari STOK_AKHIR semua kategori: produksi, minuman/barang dagang, dll.
 */
function getTransferProdukList_(ss, outlet) {
  var baseMap = buildTransferProductBaseMap_(ss);
  var hasOutlet = !!toSheetText_(outlet);

  if (hasOutlet && ss.getSheetByName(CONFIG.SHEETS.MASTER_PRODUK_OUTLET)) {
    return getTransferProdukListFromMasterOutlet_(ss, outlet, baseMap);
  }

  var result = baseMap.list.slice();
  result.sort(function(a, b) { return (a.nama || '').localeCompare(b.nama || ''); });
  return result;
}

function buildTransferProductBaseMap_(ss) {
  var rows = readStockItems_(ss);
  var productMap = {};
  var productOrder = [];

  for (var i = 0; i < rows.length; i++) {
    var item = rows[i];
    // v14.1: jangan batasi hanya kategori Produksi.
    // Barang siap jual seperti minuman/snack/kerupuk yang sudah ada di STOK_AKHIR PUSAT
    // dan didaftarkan di MASTER_PRODUK_OUTLET harus bisa langsung ditransfer ke outlet.
    if (!item.id && !item.nama) continue;

    var key = makeTransferProductKey_(item.id, item.nama);
    if (!productMap[key]) {
      productMap[key] = createTransferProduct_(item.id, item.nama, item.satuan);
      productOrder.push(key);
    }

    var product = productMap[key];
    if (!product.id && item.id) product.id = item.id;
    if (!product.nama && item.nama) product.nama = item.nama;
    if ((!product.satuan || product.satuan === '-') && item.satuan) product.satuan = item.satuan;
    if (item.lastUpdate) product.lastUpdate = item.lastUpdate;
    if (item.status) product.status = item.status;

    var lokasi = item.lokasi || CONFIG.DEFAULT_LOCATION;
    var qty = parseNumber_(item.qty);
    if (normalizeKey_(lokasi) === normalizeKey_(CONFIG.DEFAULT_LOCATION)) {
      product.stokPusat += qty;
      product.stokTersedia = product.stokPusat;
    } else {
      if (!product.outletStocks[lokasi]) product.outletStocks[lokasi] = 0;
      product.outletStocks[lokasi] += qty;
    }
  }

  var result = productOrder.map(function(key) {
    var p = productMap[key];
    p.stokTersedia = p.stokPusat;
    p.outletStockList = Object.keys(p.outletStocks).sort().map(function(lokasi) {
      return { outlet: lokasi, qty: p.outletStocks[lokasi] };
    });
    return p;
  });

  var byId = {};
  var byName = {};
  result.forEach(function(p) {
    if (p.id) byId[normalizeKey_(p.id)] = p;
    if (p.nama) byName[normalizeKey_(p.nama)] = p;
  });

  return { list: result, byId: byId, byName: byName };
}

function getTransferProdukListFromMasterOutlet_(ss, outlet, baseMap) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_PRODUK_OUTLET);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getDataRange().getValues();
  var headerMap = buildHeaderMap_(data[0]);
  var ixOutlet = findHeaderIndex_(headerMap, ['outlet', 'lokasi'], 0);
  var ixId = findHeaderIndex_(headerMap, ['idproduk', 'idbarang', 'id'], 1);
  var ixNama = findHeaderIndex_(headerMap, ['namaproduk', 'namabarang', 'produk', 'nama'], 2);
  var ixSatuan = findHeaderIndex_(headerMap, ['satuan', 'satuanproduk'], 3);
  var ixAktif = findHeaderIndex_(headerMap, ['aktif', 'status'], 4);
  var ixUrutan = findHeaderIndex_(headerMap, ['urutan', 'sort', 'no'], 5);
  var ixKet = findHeaderIndex_(headerMap, ['keterangan', 'catatan'], 6);

  var outletKey = normalizeKey_(outlet);
  var seen = {};
  var result = [];

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (normalizeKey_(row[ixOutlet]) !== outletKey) continue;
    if (!isMasterOutletActive_(row[ixAktif])) continue;

    var id = asText_(row[ixId]);
    var nama = asText_(row[ixNama]);
    var satuan = asText_(row[ixSatuan]) || '-';
    var urutan = parseNumber_(row[ixUrutan]);
    var ket = asText_(row[ixKet]);
    if (!id && !nama) continue;

    var key = makeTransferProductKey_(id, nama);
    if (seen[key]) continue;
    seen[key] = true;

    var base = findTransferProductBase_(baseMap, id, nama);
    var p = base ? cloneTransferProduct_(base) : createTransferProduct_(id, nama, satuan);
    p.id = id || p.id;
    p.nama = nama || p.nama;
    p.satuan = satuan || p.satuan || '-';
    p.urutan = urutan || 9999;
    p.keterangan = ket;
    p.masterOutlet = outlet;
    p.sumber = 'MASTER_PRODUK_OUTLET';
    result.push(p);
  }

  result.sort(function(a, b) {
    var au = Number(a.urutan || 9999);
    var bu = Number(b.urutan || 9999);
    if (au !== bu) return au - bu;
    return (a.nama || '').localeCompare(b.nama || '');
  });
  return result;
}

function isMasterOutletActive_(value) {
  var v = normalizeKey_(value);
  if (!v) return true; // kosong dianggap aktif agar input awal tidak mudah gagal tampil.
  return !(v === 'n' || v === 'no' || v === 'tidak' || v === 'false' || v === '0' || v === 'nonaktif');
}

function makeTransferProductKey_(id, nama) {
  var idKey = normalizeKey_(id);
  if (idKey) return 'id:' + idKey;
  return 'nama:' + normalizeKey_(nama);
}

function createTransferProduct_(id, nama, satuan) {
  return {
    id: asText_(id),
    nama: asText_(nama),
    satuan: asText_(satuan) || '-',
    stokPusat: 0,
    stokTersedia: 0,
    outletStocks: {},
    outletStockList: [],
    lastUpdate: '',
    status: '',
    urutan: 9999,
    keterangan: '',
    masterOutlet: ''
  };
}

function cloneTransferProduct_(p) {
  var clone = createTransferProduct_(p.id, p.nama, p.satuan);
  clone.stokPusat = parseNumber_(p.stokPusat || p.stokTersedia);
  clone.stokTersedia = clone.stokPusat;
  clone.outletStocks = Object.assign({}, p.outletStocks || {});
  clone.outletStockList = (p.outletStockList || []).slice();
  clone.lastUpdate = p.lastUpdate || '';
  clone.status = p.status || '';
  clone.urutan = p.urutan || 9999;
  clone.keterangan = p.keterangan || '';
  clone.masterOutlet = p.masterOutlet || '';
  return clone;
}

function findTransferProductBase_(baseMap, id, nama) {
  var idKey = normalizeKey_(id);
  var nameKey = normalizeKey_(nama);
  if (idKey && baseMap.byId[idKey]) return baseMap.byId[idKey];
  if (nameKey && baseMap.byName[nameKey]) return baseMap.byName[nameKey];
  return null;
}

/**
 * MASTER_BARANG_PRODUKSI
 * Kolom A: ID Barang Produksi / Lauk
 * Kolom B: Produksi / Nama Lauk
 * Kolom C: Bahan Baku yang dipakai
 * Kolom D: Satuan Bahan
 * Kolom E: Satuan Produksi
 *
 * Revisi v2:
 * - idProduksi disimpan di level grup produksi.
 * - idBahan dicari dari STOK_AKHIR berdasarkan nama bahan baku agar bisa dikurangi sebagai pengeluaran.
 */
function handleGetBarangProduksi_(ss, requestData) {
  if (requestData && requestData._user && !profileCanReadCategory_(requestData._user, 'produksi')) {
    return dataResponse_({ success: false, message: 'Akses Produksi ditolak untuk user login.' });
  }

  var sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_PRODUKSI);
  if (!sheet) return dataResponse_({ success: false, message: 'Sheet MASTER_BARANG_PRODUKSI tidak ditemukan.' });

  var stokMaps = buildStockMaps_(ss);
  var data = sheet.getDataRange().getValues();
  var grouped = {};

  for (var i = 1; i < data.length; i++) {
    var idProduksi = asText_(data[i][0]);
    var produksi = asText_(data[i][1]);
    var bahanBaku = asText_(data[i][2]);
    var satuanBahanMaster = asText_(data[i][3]);
    var satuanProduksi = asText_(data[i][4]) || '-';

    if (!produksi || !bahanBaku) continue;

    var produksiKey = normalizeKey_(produksi);
    var bahanStock = getStockItem_(stokMaps, CONFIG.DEFAULT_LOCATION, '', bahanBaku);

    if (!grouped[produksiKey]) {
      grouped[produksiKey] = {
        idProduksi: idProduksi,
        produksi: produksi,
        satuanProduksi: satuanProduksi,
        bahan: []
      };
    }

    // Jika baris awal kosong tapi baris berikutnya ada ID, tetap pakai ID produksi yang ketemu.
    if (!grouped[produksiKey].idProduksi && idProduksi) grouped[produksiKey].idProduksi = idProduksi;
    if (!grouped[produksiKey].satuanProduksi && satuanProduksi) grouped[produksiKey].satuanProduksi = satuanProduksi;

    grouped[produksiKey].bahan.push({
      idBahan: bahanStock ? bahanStock.id : '',
      bahanBaku: bahanBaku,
      satuanBahan: satuanBahanMaster || (bahanStock ? bahanStock.satuan : '-') || '-',
      stokTersedia: bahanStock ? bahanStock.qty : 0,
      // Tambahan flag agar frontend bisa memperingatkan jika nama bahan tidak cocok dengan STOK_AKHIR.
      stokMatch: !!bahanStock
    });
  }

  var result = Object.keys(grouped).map(function(key) { return grouped[key]; });
  result.sort(function(a, b) { return a.produksi.localeCompare(b.produksi); });

  return dataResponse_({ success: true, data: result });
}


/**
 * BARU v3: Data stok produksi untuk halaman Lihat Stok.
 * Kenapa dibuat endpoint khusus?
 * - MASTER_BARANG_PRODUKSI berisi banyak baris bahan untuk 1 produksi/lauk.
 * - Kalau dipanggil lewat getBarang biasa, hasilnya akan duplikat.
 * - Endpoint ini mengubahnya menjadi 1 baris unik per ID/Nama Produksi.
 *
 * Prioritas qty stok:
 * 1) Ambil dari STOK_AKHIR berdasarkan ID Produksi atau Nama Produksi.
 * 2) Jika belum ada di STOK_AKHIR, fallback ke total TRANSAKSI_OUTPUT_PRODUKSI.
 *    Fallback ini membantu masa transisi, tapi tetap idealnya STOK_AKHIR yang jadi sumber utama.
 */
function handleGetStokProduksiReport_(ss, requestData) {
  if (requestData && requestData._user && !profileCanReadCategory_(requestData._user, 'produksi')) {
    return dataResponse_({ success: false, message: 'Akses Lihat Stok Produksi ditolak untuk user login.' });
  }

  requestData = requestData || {};
  requestData.category = 'Produksi';
  var fromStok = handleGetStokAkhirReportData_(ss, requestData);
  if (fromStok.length) {
    return dataResponse_({ success: true, data: fromStok });
  }

  // Fallback jika STOK_AKHIR belum terisi produksi: pakai log produksi lama.
  var sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_PRODUKSI);
  if (!sheet) return dataResponse_({ success: false, message: 'Sheet MASTER_BARANG_PRODUKSI tidak ditemukan.' });

  var stokMaps = buildStockMaps_(ss);
  var produksiSums = buildProduksiSumMaps_(ss);
  var data = sheet.getDataRange().getValues();
  var unique = {};

  for (var i = 1; i < data.length; i++) {
    var idProduksi = asText_(data[i][0]);
    var produksi = asText_(data[i][1]);
    var satuanProduksi = asText_(data[i][4]);
    if (!produksi) continue;

    var key = idProduksi ? 'id:' + normalizeKey_(idProduksi) : 'nama:' + normalizeKey_(produksi);
    if (unique[key]) continue;

    var stokItem = getStockItem_(stokMaps, CONFIG.DEFAULT_LOCATION, idProduksi, produksi);
    var sumItem = null;
    if (idProduksi) sumItem = produksiSums.byId[normalizeKey_(idProduksi)] || null;
    if (!sumItem) sumItem = produksiSums.byName[normalizeKey_(produksi)] || null;

    unique[key] = {
      lokasi: CONFIG.DEFAULT_LOCATION,
      kategori: 'Produksi',
      id: idProduksi || (stokItem ? stokItem.id : ''),
      nama: produksi,
      satuanStok: (stokItem && stokItem.satuan) || satuanProduksi || (sumItem ? sumItem.satuan : '-') || '-',
      stokTersedia: stokItem ? stokItem.qty : (sumItem ? sumItem.qty : 0),
      lastUpdate: (stokItem && stokItem.lastUpdate) || (sumItem ? sumItem.lastUpdate : ''),
      status: (stokItem && stokItem.status) || '',
      sumber: stokItem ? 'STOK_AKHIR' : (sumItem ? 'TRANSAKSI_OUTPUT_PRODUKSI' : 'MASTER_BARANG_PRODUKSI')
    };
  }

  var result = Object.keys(unique).map(function(key) { return unique[key]; });
  result.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
  return dataResponse_({ success: true, data: result });
}

/** v8B: Laporan langsung dari STOK_AKHIR header baru, sudah mendukung Lokasi. */
function handleGetStokAkhirReport_(ss, requestData) {
  var category = normalizeCategoryRequest_((requestData || {}).category || (requestData || {}).kategori || (requestData || {}).sheetName || '');
  var permissionKey = categoryPermissionKey_(category);
  if (requestData && requestData._user && !permissionKey && !profileHasPermission_(requestData._user, 'lihatStok')) {
    return dataResponse_({ success: false, message: 'Akses laporan semua stok ditolak untuk user login.' });
  }
  if (requestData && requestData._user && permissionKey && !profileCanReadCategory_(requestData._user, permissionKey)) {
    return dataResponse_({ success: false, message: 'Akses kategori stok ini ditolak untuk user login.' });
  }
  return dataResponse_({ success: true, data: handleGetStokAkhirReportData_(ss, requestData || {}) });
}

function handleGetStokAkhirReportData_(ss, requestData) {
  var category = normalizeCategoryRequest_(requestData.category || requestData.kategori || requestData.sheetName || '');
  var lokasi = asText_(requestData.lokasi || requestData.location || '');
  var maps = buildStockMaps_(ss);
  var hiddenParentIds = getBahanMentahParentIdMap_(ss);
  var rows = [];

  for (var i = 0; i < maps.allRows.length; i++) {
    var item = maps.allRows[i];
    if (item.id && hiddenParentIds[normalizeKey_(item.id)]) continue;
    var okCategory = category ? categoryMatches_(item.kategori, category) : true;
    var okLokasi = lokasi ? normalizeKey_(item.lokasi) === normalizeKey_(lokasi) : true;
    if (!okCategory || !okLokasi) continue;
    rows.push({
      lokasi: item.lokasi,
      kategori: displayCategoryLabel_(item.kategori),
      id: item.id,
      nama: item.nama,
      satuanStok: item.satuan,
      stokTersedia: item.qty,
      lastUpdate: item.lastUpdate,
      status: item.status,
      sumber: 'STOK_AKHIR'
    });
  }

  rows.sort(function(a, b) {
    var byLoc = a.lokasi.localeCompare(b.lokasi);
    if (byLoc !== 0) return byLoc;
    return a.nama.localeCompare(b.nama);
  });
  return rows;
}

function normalizeCategoryRequest_(value) {
  var key = normalizeKey_(value).replace(/\s+/g, '');
  var map = {
    'master_bahan_mentah': 'Bahan Mentah',
    'masterbahanmentah': 'Bahan Mentah',
    'bahanmentah': 'Bahan Mentah',
    'master_bahan_baku': 'Bahan Baku',
    'masterbahanbaku': 'Bahan Baku',
    'bahanbaku': 'Bahan Baku',
    'master_kemasan': 'Kemasan',
    'masterkemasan': 'Kemasan',
    'kemasan': 'Kemasan',
    'master_minuman': 'Minuman & Barang Dagang',
    'masterminuman': 'Minuman & Barang Dagang',
    'minuman': 'Minuman & Barang Dagang',
    'minumanbarangdagang': 'Minuman & Barang Dagang',
    'master_kebersihan': 'Kebersihan',
    'masterkebersihan': 'Kebersihan',
    'kebersihan': 'Kebersihan',
    'produksi': 'Produksi'
  };
  return map[key] || asText_(value);
}

function displayCategoryLabel_(value) {
  var key = normalizeKey_(value);
  if (key === 'minuman' || key === 'minumanbarangdagang') return 'Minuman & Barang Dagang';
  return asText_(value);
}

function categoryMatches_(left, right) {
  var a = normalizeKey_(left);
  var b = normalizeKey_(right);
  if (a === b) return true;
  var aMinuman = (a === 'minuman' || a === 'minumanbarangdagang');
  var bMinuman = (b === 'minuman' || b === 'minumanbarangdagang');
  return aMinuman && bMinuman;
}

function isValidDateText_(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(asText_(value));
}

function categoryAccessError_(profile, kategori) {
  var permissionKey = categoryPermissionKey_(kategori);
  if (!permissionKey) return 'Kategori "' + asText_(kategori) + '" tidak dikenali.';
  if (!profileHasPermission_(profile, permissionKey)) {
    return 'Akses kategori "' + displayCategoryLabel_(kategori) + '" ditolak untuk user login.';
  }
  return '';
}

/**
 * REVISI v5: Simpan input stok biasa + preparasi ayam.
 * Payload:
 * - inputRows  : baris masuk ke TRANSAKSI_INPUT_STOK.
 * - outputRows : opsional, baris keluar ke TRANSAKSI_OUTPUT_STOK untuk parent yang diproses/preparasi.
 *
 * Untuk ayam parent:
 * 1) Ayam Potong 8/11 dicatat masuk.
 * 2) Ayam Potong 8/11 dicatat keluar dengan tujuan Preparasi.
 * 3) Ayam 8/11 Kremes/Premix dicatat masuk sebagai stok operasional siap produksi.
 */
function handleSimpanInputStokPreparasi_(ss, requestData) {
  var profile = requestData._user;
  if (!profile || !profile.nama) return dataResponse_({ success: false, message: 'Petugas tidak terbaca. Silakan login ulang.' });
  var inputRows = requestData.inputRows || [];
  var outputRows = requestData.outputRows || [];

  if (!inputRows.length) return dataResponse_({ success: false, message: 'Tidak ada data input stok untuk disimpan.' });

  for (var i = 0; i < inputRows.length; i++) {
    var row = inputRows[i] || [];
    if (!Array.isArray(row)) return dataResponse_({ success: false, message: 'Format baris input stok tidak valid.' });
    if (!isValidDateText_(row[0])) return dataResponse_({ success: false, message: 'Tanggal input stok tidak valid.' });
    if (!asText_(row[3]) && !asText_(row[4])) return dataResponse_({ success: false, message: 'ID/Nama barang input wajib diisi.' });
    if (parseNumber_(row[5]) <= 0 || parseNumber_(row[7]) <= 0) return dataResponse_({ success: false, message: 'Qty input stok wajib lebih dari 0.' });
    var accessError = categoryAccessError_(profile, row[2]);
    if (accessError) return dataResponse_({ success: false, message: accessError });
    row[1] = profile.nama;
  }

  for (var j = 0; j < outputRows.length; j++) {
    var outRow = outputRows[j] || [];
    if (!Array.isArray(outRow)) return dataResponse_({ success: false, message: 'Format baris preparasi tidak valid.' });
    if (!isValidDateText_(outRow[0])) return dataResponse_({ success: false, message: 'Tanggal preparasi tidak valid.' });
    if (!asText_(outRow[3]) && !asText_(outRow[4])) return dataResponse_({ success: false, message: 'ID/Nama barang preparasi wajib diisi.' });
    if (parseNumber_(outRow[5]) <= 0) return dataResponse_({ success: false, message: 'Qty preparasi wajib lebih dari 0.' });
    var outAccessError = categoryAccessError_(profile, outRow[2]);
    if (outAccessError) return dataResponse_({ success: false, message: outAccessError });
    outRow[1] = profile.nama;
  }

  return withScriptLock_(function() {
    var txId = buildTxId_('IN');
    var inputSheet = getOrCreateSheet_(ss, CONFIG.SHEETS.INPUT_STOK, HEADER_INPUT_STOK, false);

    var rowsInput = inputRows.map(function(row) {
      return [txId].concat(row.map(toSheetText_));
    });

    inputSheet.getRange(inputSheet.getLastRow() + 1, 1, rowsInput.length, HEADER_INPUT_STOK.length).setValues(rowsInput);

    if (outputRows.length) {
      var outputSheet = getOrCreateSheet_(ss, CONFIG.SHEETS.OUTPUT_STOK, HEADER_OUTPUT_STOK, false);
      var rowsOutput = outputRows.map(function(row) {
        return [txId].concat(row.map(toSheetText_));
      });
      outputSheet.getRange(outputSheet.getLastRow() + 1, 1, rowsOutput.length, HEADER_OUTPUT_STOK.length).setValues(rowsOutput);
    }

    appendAuditLog_(ss, requestData, 'SUCCESS', 'Input stok disimpan. ID: ' + txId + '. Baris input: ' + rowsInput.length + ', preparasi: ' + outputRows.length + '.');
    return dataResponse_({
      success: true,
      message: outputRows.length
        ? 'Input stok dan preparasi berhasil disimpan dengan ID: ' + txId
        : 'Input stok berhasil disimpan dengan ID: ' + txId
    });
  });
}


/** v14.8: daftar barang input stok berdasarkan tanggal untuk popup cetak input stok 58mm.
 * - Jika tanggal kosong: tetap mengembalikan tanggal yang punya data.
 * - Jika tanggal dipilih dan tidak ada barang: data kosong, supaya dropdown Nama Barang kosong.
 */
function handleGetInputPrintOptions_(ss, requestData) {
  if (requestData && requestData._user && !profileHasAnyStockCategory_(requestData._user)) {
    return dataResponse_({ success: false, message: 'Akses cetak input stok ditolak untuk user login.' });
  }

  requestData = requestData || {};
  var tanggalFilter = asText_(requestData.tanggal);
  var sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT_STOK);
  if (!sheet || sheet.getLastRow() < 2) {
    return dataResponse_({ success: true, dates: [], data: [] });
  }

  var values = sheet.getDataRange().getValues();
  var hm = buildHeaderMap_(values[0]);
  var ixTanggal = findHeaderIndex_(hm, ['tanggal'], 1);
  var ixId = findHeaderIndex_(hm, ['id barang', 'id produk', 'kode barang'], 4);
  var ixNama = findHeaderIndex_(hm, ['nama barang', 'nama produk'], 5);
  var ixQtyStok = findHeaderIndex_(hm, ['qty stok', 'stok masuk', 'qty masuk'], 8);
  var ixSatuan = findHeaderIndex_(hm, ['satuan stok', 'satuan'], 9);
  var ixKategori = findHeaderIndex_(hm, ['kategori'], 3);

  var dateSeen = {};
  var dates = [];
  var map = {};

  for (var r = 1; r < values.length; r++) {
    var tgl = formatSheetDate_(values[r][ixTanggal]);
    if (!tgl) continue;
    if (!dateSeen[tgl]) {
      dateSeen[tgl] = true;
      dates.push(tgl);
    }

    if (tanggalFilter && tgl !== tanggalFilter) continue;

    var id = asText_(values[r][ixId]);
    var nama = asText_(values[r][ixNama]);
    if (!id && !nama) continue;

    var key = normalizeKey_(id || nama);
    if (!map[key]) {
      map[key] = {
        key: key,
        id: id,
        nama: nama,
        satuan: asText_(values[r][ixSatuan]) || '-',
        kategori: displayCategoryLabel_(values[r][ixKategori]),
        qtyMasuk: 0
      };
    }
    map[key].qtyMasuk += parseNumber_(values[r][ixQtyStok]);
    if (!map[key].id) map[key].id = id;
    if (!map[key].nama) map[key].nama = nama;
    if (!map[key].satuan || map[key].satuan === '-') map[key].satuan = asText_(values[r][ixSatuan]) || '-';
    if (!map[key].kategori) map[key].kategori = displayCategoryLabel_(values[r][ixKategori]);
  }

  var out = Object.keys(map).map(function(k) { return map[k]; })
    .filter(function(item) { return parseNumber_(item.qtyMasuk) > 0; })
    .sort(function(a, b) { return (a.nama || '').localeCompare(b.nama || ''); });

  dates.sort();
  return dataResponse_({ success: true, dates: dates, data: out });
}

/** v14.8: ambil agregat stok masuk per tanggal untuk beberapa barang sekaligus. */
function handleGetInputPrintData_(ss, requestData) {
  if (requestData && requestData._user && !profileHasAnyStockCategory_(requestData._user)) {
    return dataResponse_({ success: false, message: 'Akses cetak input stok ditolak untuk user login.' });
  }

  requestData = requestData || {};
  var tanggal = asText_(requestData.tanggal);
  var barangKeys = requestData.barangKeys || [];
  var printAll = requestData.printAll === true || asText_(requestData.barangKey).toUpperCase() === '__ALL__';

  if (!Array.isArray(barangKeys)) barangKeys = [barangKeys];
  if (!barangKeys.length && requestData.barangKey) barangKeys = [requestData.barangKey];
  barangKeys = barangKeys.map(function(k) { return normalizeKey_(k); }).filter(String);

  var keyFilter = {};
  barangKeys.forEach(function(k) { keyFilter[k] = true; });

  if (!tanggal || (!printAll && !barangKeys.length)) {
    return dataResponse_({ success: false, message: 'Tanggal dan minimal 1 nama barang wajib dipilih.' });
  }

  var sheet = ss.getSheetByName(CONFIG.SHEETS.INPUT_STOK);
  if (!sheet || sheet.getLastRow() < 2) {
    return dataResponse_({ success: false, message: 'Data input stok masih kosong.' });
  }

  var values = sheet.getDataRange().getValues();
  var hm = buildHeaderMap_(values[0]);
  var ixTx = findHeaderIndex_(hm, ['id transaksi'], 0);
  var ixTanggal = findHeaderIndex_(hm, ['tanggal'], 1);
  var ixPetugas = findHeaderIndex_(hm, ['petugas'], 2);
  var ixKategori = findHeaderIndex_(hm, ['kategori'], 3);
  var ixId = findHeaderIndex_(hm, ['id barang', 'id produk', 'kode barang'], 4);
  var ixNama = findHeaderIndex_(hm, ['nama barang', 'nama produk'], 5);
  var ixQtyStok = findHeaderIndex_(hm, ['qty stok', 'stok masuk', 'qty masuk'], 8);
  var ixSatuanStok = findHeaderIndex_(hm, ['satuan stok', 'satuan'], 9);

  var map = {};

  for (var r = 1; r < values.length; r++) {
    var rowTanggal = formatSheetDate_(values[r][ixTanggal]);
    if (rowTanggal !== tanggal) continue;

    var rowId = asText_(values[r][ixId]);
    var rowNama = asText_(values[r][ixNama]);
    if (!rowId && !rowNama) continue;

    var rowKeyId = normalizeKey_(rowId);
    var rowKeyNama = normalizeKey_(rowNama);
    var rowKey = rowKeyId || rowKeyNama;
    if (!printAll && !keyFilter[rowKeyId] && !keyFilter[rowKeyNama]) continue;

    if (!map[rowKey]) {
      map[rowKey] = {
        key: rowKey,
        tanggal: tanggal,
        idBarang: rowId,
        namaBarang: rowNama,
        qtyMasuk: 0,
        satuan: asText_(values[r][ixSatuanStok]) || '-',
        kategori: displayCategoryLabel_(values[r][ixKategori]),
        petugasList: [],
        txList: [],
        petugasSeen: {},
        txSeen: {}
      };
    }

    var item = map[rowKey];
    item.qtyMasuk += parseNumber_(values[r][ixQtyStok]);
    if (!item.idBarang) item.idBarang = rowId;
    if (!item.namaBarang) item.namaBarang = rowNama;
    if (!item.satuan || item.satuan === '-') item.satuan = asText_(values[r][ixSatuanStok]) || '-';
    if (!item.kategori) item.kategori = displayCategoryLabel_(values[r][ixKategori]);

    var tx = asText_(values[r][ixTx]);
    if (tx && !item.txSeen[tx]) { item.txSeen[tx] = true; item.txList.push(tx); }
    var petugas = asText_(values[r][ixPetugas]);
    if (petugas && !item.petugasSeen[petugas]) { item.petugasSeen[petugas] = true; item.petugasList.push(petugas); }
  }

  var items = Object.keys(map).map(function(k) {
    var item = map[k];
    return {
      tanggal: item.tanggal,
      idBarang: item.idBarang,
      namaBarang: item.namaBarang,
      qtyMasuk: item.qtyMasuk,
      satuan: item.satuan || '-',
      kategori: item.kategori,
      petugas: item.petugasList.join(', '),
      idTransaksi: item.txList.join(', '),
      jumlahTransaksi: item.txList.length
    };
  }).filter(function(item) {
    return parseNumber_(item.qtyMasuk) > 0;
  });

  // Kalau user menambahkan barang dengan urutan tertentu, ikuti urutan pilihan user.
  if (!printAll && barangKeys.length) {
    var order = {};
    barangKeys.forEach(function(k, idx) { order[k] = idx; });
    items.sort(function(a, b) {
      var ka = normalizeKey_(a.idBarang) || normalizeKey_(a.namaBarang);
      var kb = normalizeKey_(b.idBarang) || normalizeKey_(b.namaBarang);
      var oa = Object.prototype.hasOwnProperty.call(order, ka) ? order[ka] : 9999;
      var ob = Object.prototype.hasOwnProperty.call(order, kb) ? order[kb] : 9999;
      if (oa !== ob) return oa - ob;
      return (a.namaBarang || '').localeCompare(b.namaBarang || '');
    });
  } else {
    items.sort(function(a, b) { return (a.namaBarang || '').localeCompare(b.namaBarang || ''); });
  }

  if (!items.length) {
    return dataResponse_({ success: false, message: 'Data stok masuk untuk tanggal dan barang ini tidak ditemukan.' });
  }

  var first = items[0];
  return dataResponse_({
    success: true,
    data: {
      tanggal: tanggal,
      items: items,
      jumlahBarang: items.length,
      // field lama tetap disediakan agar kompatibel dengan format single barang.
      idBarang: first.idBarang,
      namaBarang: first.namaBarang,
      qtyMasuk: first.qtyMasuk,
      satuan: first.satuan,
      kategori: first.kategori,
      petugas: first.petugas,
      idTransaksi: first.idTransaksi,
      jumlahTransaksi: first.jumlahTransaksi
    }
  });
}

/** GENERAL: simpan transaksi lama, tetapi dibuat lebih aman dan bulk write. */
function handleSimpanTransaksi_(ss, requestData) {
  var profile = requestData._user;
  if (!profile || !profile.nama) return dataResponse_({ success: false, message: 'Petugas tidak terbaca. Silakan login ulang.' });
  var destSheetName = requestData.sheetName;
  var sheetData = requestData.data || [];

  if (!sheetData.length) return dataResponse_({ success: false, message: 'Tidak ada data untuk disimpan.' });
  if (destSheetName !== CONFIG.SHEETS.OUTPUT_STOK && destSheetName !== CONFIG.SHEETS.STOK_OPNAME) {
    return dataResponse_({ success: false, message: 'Sheet tujuan tidak diizinkan untuk endpoint simpanTransaksi.' });
  }

  var header = null;
  if (destSheetName === CONFIG.SHEETS.OUTPUT_STOK) header = HEADER_OUTPUT_STOK;

  var destSheet = header ? getOrCreateSheet_(ss, destSheetName, header, false) : ss.getSheetByName(destSheetName);
  if (!destSheet) return dataResponse_({ success: false, message: "Sheet tujuan '" + destSheetName + "' tidak ditemukan." });

  if (destSheetName === CONFIG.SHEETS.OUTPUT_STOK) {
    var stokMaps = buildStockMaps_(ss);
    var requested = {};
    for (var i = 0; i < sheetData.length; i++) {
      var outRow = sheetData[i] || [];
      if (!Array.isArray(outRow)) return dataResponse_({ success: false, message: 'Format baris output stok tidak valid.' });
      if (!isValidDateText_(outRow[0])) return dataResponse_({ success: false, message: 'Tanggal output stok tidak valid.' });
      var outputAccessError = categoryAccessError_(profile, outRow[2]);
      if (outputAccessError) return dataResponse_({ success: false, message: outputAccessError });
      if (categoryPermissionKey_(outRow[2]) === 'produksi') {
        return dataResponse_({ success: false, message: 'Output Produksi wajib memakai endpoint Produksi khusus.' });
      }
      if (!asText_(outRow[3]) && !asText_(outRow[4])) return dataResponse_({ success: false, message: 'ID/Nama barang output wajib diisi.' });
      var qtyKeluar = parseNumber_(outRow[5]);
      if (qtyKeluar <= 0) return dataResponse_({ success: false, message: 'Qty keluar wajib lebih dari 0.' });
      var tujuan = asText_(outRow[7]);
      if (normalizeKey_(tujuan) === 'transfer' && (!asText_(outRow[8]) || !asText_(outRow[9]))) {
        return dataResponse_({ success: false, message: 'Tujuan Transfer wajib isi Outlet dan PIC.' });
      }
      var stock = getStockItem_(stokMaps, CONFIG.DEFAULT_LOCATION, outRow[3], outRow[4]);
      var stockQty = stock ? parseNumber_(stock.qty) : 0;
      var key = productKey_(outRow[3], outRow[4]);
      if (!requested[key]) requested[key] = { qty: 0, stokQty: stockQty, nama: asText_(outRow[4] || outRow[3]), satuan: (stock && stock.satuan) || asText_(outRow[6]) };
      requested[key].qty += qtyKeluar;
      outRow[1] = profile.nama;
      if (stock && stock.satuan) outRow[6] = stock.satuan;
    }
    for (var reqKey in requested) {
      var req = requested[reqKey];
      if (req.qty > req.stokQty) {
        return dataResponse_({ success: false, message: 'Qty keluar ' + req.nama + ' melebihi stok PUSAT. Stok tersedia: ' + req.stokQty + ' ' + (req.satuan || '') + '.' });
      }
    }
  }

  if (destSheetName === CONFIG.SHEETS.STOK_OPNAME) {
    if (!profileHasPermission_(profile, 'opname')) {
      return dataResponse_({ success: false, message: 'Akses Stok Opname ditolak untuk user login.' });
    }
    var opnameStockMaps = buildStockMaps_(ss);
    for (var j = 0; j < sheetData.length; j++) {
      var opnRow = sheetData[j] || [];
      if (!Array.isArray(opnRow)) return dataResponse_({ success: false, message: 'Format baris stok opname tidak valid.' });
      if (!isValidDateText_(opnRow[0])) return dataResponse_({ success: false, message: 'Tanggal stok opname tidak valid.' });
      var opnameAccessError = categoryAccessError_(profile, opnRow[2]);
      if (opnameAccessError) return dataResponse_({ success: false, message: opnameAccessError });
      if (!asText_(opnRow[3]) && !asText_(opnRow[4])) return dataResponse_({ success: false, message: 'ID/Nama barang opname wajib diisi.' });
      var stokFisik = parseNumber_(opnRow[6]);
      if (stokFisik < 0) return dataResponse_({ success: false, message: 'Stok fisik tidak boleh minus.' });
      var opnameStock = getStockItem_(opnameStockMaps, CONFIG.DEFAULT_LOCATION, opnRow[3], opnRow[4]);
      var stokSistem = opnameStock ? parseNumber_(opnameStock.qty) : 0;
      opnRow[1] = profile.nama;
      opnRow[5] = stokSistem;
      opnRow[7] = stokFisik - stokSistem;
      if (opnameStock && opnameStock.satuan) opnRow[8] = opnameStock.satuan;
    }
  }

  return withScriptLock_(function() {
    var prefix = destSheetName === CONFIG.SHEETS.OUTPUT_STOK ? 'OUT' : 'OPN';
    var txId = buildTxId_(prefix);
    var rows = sheetData.map(function(row) {
      return [txId].concat(row.map(toSheetText_));
    });

    destSheet.getRange(destSheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    appendAuditLog_(ss, requestData, 'SUCCESS', 'Transaksi ' + destSheetName + ' disimpan. ID: ' + txId + '. Baris: ' + rows.length + '.');
    return dataResponse_({ success: true, message: 'Berhasil disimpan dengan ID: ' + txId });
  });
}

/**
 * REVISI v2: Simpan produksi jadi dan pemakaian bahan sekaligus.
 *
 * 1) TRANSAKSI_OUTPUT_PRODUKSI = stok lauk/jadi bertambah untuk dijual.
 *    Format: [ID Transaksi, Tanggal, Petugas, ID Barang Produksi, Produksi, Qty Produksi, Satuan Produksi]
 *
 * 2) TRANSAKSI_OUTPUT_STOK = bahan mentah/baku berkurang sebagai pengeluaran produksi.
 *    Format: [ID Transaksi, Tanggal, Petugas, Kategori, ID Barang, Nama Barang, Qty Keluar, Satuan, Tujuan, Outlet, PIC]
 */
function handleSimpanOutputProduksiV2_(ss, requestData) {
  var profile = requestData._user;
  if (!profile || !profile.nama) return dataResponse_({ success: false, message: 'Petugas tidak terbaca. Silakan login ulang.' });
  var produksiRows = requestData.produksiRows || [];
  var bahanRows = requestData.bahanRows || [];

  if (!produksiRows.length) return dataResponse_({ success: false, message: 'Tidak ada data produksi/lauk untuk disimpan.' });
  if (!bahanRows.length) return dataResponse_({ success: false, message: 'Tidak ada data pemakaian bahan untuk disimpan.' });

  // REVISI v6: validasi backend. Walaupun menu Produksi disembunyikan di frontend,
  // server tetap menolak simpan produksi jika petugas tidak punya Produksi = Y.
  if (!profileHasPermission_(profile, 'produksi')) {
    return dataResponse_({ success: false, message: 'Akses ditolak. Role petugas ini tidak memiliki izin Produksi.' });
  }

  for (var i = 0; i < produksiRows.length; i++) {
    if (!isValidDateText_(produksiRows[i].tanggal)) return dataResponse_({ success: false, message: 'Tanggal produksi tidak valid.' });
    produksiRows[i].petugas = profile.nama;
  }
  for (var j = 0; j < bahanRows.length; j++) {
    if (!isValidDateText_(bahanRows[j].tanggal)) return dataResponse_({ success: false, message: 'Tanggal pemakaian bahan tidak valid.' });
    bahanRows[j].petugas = profile.nama;
  }

  return withScriptLock_(function() {
    var txId = buildTxId_('PRD');

    var produksiSheet = getOrCreateSheet_(ss, CONFIG.SHEETS.OUTPUT_PRODUKSI, HEADER_OUTPUT_PRODUKSI, true);
    var outputSheet = getOrCreateSheet_(ss, CONFIG.SHEETS.OUTPUT_STOK, HEADER_OUTPUT_STOK, false);

    var rowsProduksi = produksiRows.map(function(row) {
      return [
        txId,
        toSheetText_(row.tanggal),
        toSheetText_(row.petugas),
        toSheetText_(row.idProduksi),
        toSheetText_(row.produksi),
        toSheetText_(row.qtyProduksi),
        toSheetText_(row.satuanProduksi)
      ];
    });

    var rowsBahan = bahanRows.map(function(row) {
      return [
        txId,
        toSheetText_(row.tanggal),
        toSheetText_(row.petugas),
        'Produksi',                 // Kategori pengeluaran bahan.
        toSheetText_(row.idBahan),
        toSheetText_(row.bahanBaku),
        toSheetText_(row.qtyKeluar),
        toSheetText_(row.satuanBahan),
        'Produksi',                 // Tujuan khusus agar jelas ini bukan transfer outlet.
        '',                         // Outlet kosong.
        ''                          // PIC kosong.
      ];
    });

    produksiSheet.getRange(produksiSheet.getLastRow() + 1, 1, rowsProduksi.length, HEADER_OUTPUT_PRODUKSI.length).setValues(rowsProduksi);
    outputSheet.getRange(outputSheet.getLastRow() + 1, 1, rowsBahan.length, HEADER_OUTPUT_STOK.length).setValues(rowsBahan);

    appendAuditLog_(ss, requestData, 'SUCCESS', 'Produksi disimpan. ID: ' + txId + '. Produksi: ' + rowsProduksi.length + ', bahan: ' + rowsBahan.length + '.');
    return dataResponse_({
      success: true,
      message: 'Produksi jadi dan pemakaian bahan berhasil disimpan dengan ID: ' + txId
    });
  });
}


/**
 * REVISI v14.1: Simpan Transfer Produk Outlet.
 * Format sheet:
 * ID Transfer | Tanggal | Petugas | Outlet | PIC Penerima | ID Produk | Nama Produk | Qty Transfer | Satuan | Catatan
 */
function handleSimpanTransferProduksi_(ss, requestData) {
  var profile = requestData._user;
  if (!profile || !profile.nama) return dataResponse_({ success: false, message: 'Petugas tidak terbaca. Silakan login ulang.' });
  var tanggal = toSheetText_(requestData.tanggal);
  var petugas = toSheetText_(profile.nama);
  var outlet = toSheetText_(requestData.outlet);
  var pic = toSheetText_(requestData.pic || requestData.picPenerima);
  var rows = requestData.rows || requestData.data || [];

  if (!tanggal) return dataResponse_({ success: false, message: 'Tanggal transfer wajib diisi.' });
  if (!isValidDateText_(tanggal)) return dataResponse_({ success: false, message: 'Tanggal transfer tidak valid.' });
  if (!petugas) return dataResponse_({ success: false, message: 'Petugas tidak terbaca. Silakan logout lalu login ulang.' });
  if (!outlet) return dataResponse_({ success: false, message: 'Outlet tujuan wajib dipilih.' });
  if (!pic) return dataResponse_({ success: false, message: 'PIC penerima wajib dipilih.' });
  if (!rows.length) return dataResponse_({ success: false, message: 'Tidak ada produk yang ditransfer.' });

  if (!profileHasPermission_(profile, 'transferProduksi') && !profileHasPermission_(profile, 'produksi')) {
    return dataResponse_({ success: false, message: 'Akses ditolak. Role petugas ini tidak memiliki izin Transfer Produk Outlet.' });
  }

  var stokMaps = buildStockMaps_(ss);
  var txId = buildTxId_('TFP');
  var output = [];
  var requestedByProduct = {};

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i] || {};
    var idProduk = toSheetText_(r.idProduk || r.id || r[0]);
    var namaProduk = toSheetText_(r.namaProduk || r.nama || r[1]);
    var qtyTransfer = parseNumber_(r.qtyTransfer || r.qty || r[2]);
    var satuan = toSheetText_(r.satuan || r[3]) || '-';
    var catatan = toSheetText_(r.catatan || r.keterangan || r[4]);

    if (!idProduk && !namaProduk) continue;
    if (qtyTransfer < 0) return dataResponse_({ success: false, message: 'Qty transfer tidak boleh minus untuk ' + (namaProduk || idProduk) + '.' });
    // v12: qty kosong / 0 boleh dilewati. Ini supaya tabel bisa menampilkan banyak produk,
    // tapi user cukup isi menu yang memang dikirim.
    if (qtyTransfer <= 0) continue;

    var stok = getStockItem_(stokMaps, CONFIG.DEFAULT_LOCATION, idProduk, namaProduk);
    var stokQty = stok ? parseNumber_(stok.qty) : 0;
    var productKey = normalizeKey_(idProduk || namaProduk);
    if (!requestedByProduct[productKey]) {
      requestedByProduct[productKey] = { qty: 0, stokQty: stokQty, nama: namaProduk || idProduk, satuan: (stok && stok.satuan) || satuan };
    }
    requestedByProduct[productKey].qty += qtyTransfer;

    output.push([
      txId, tanggal, petugas, outlet, pic, idProduk, namaProduk, toSheetText_(qtyTransfer), satuan, catatan
    ]);
  }

  if (!output.length) return dataResponse_({ success: false, message: 'Tidak ada qty transfer yang diisi.' });

  for (var key in requestedByProduct) {
    var req = requestedByProduct[key];
    if (req.qty > req.stokQty) {
      return dataResponse_({
        success: false,
        message: 'Total qty transfer ' + req.nama + ' melebihi stok PUSAT. Stok tersedia: ' + req.stokQty + ' ' + req.satuan + '.'
      });
    }
  }

  return withScriptLock_(function() {
    var sheet = getOrCreateSheet_(ss, CONFIG.SHEETS.TRANSFER_PRODUKSI, HEADER_TRANSFER_PRODUKSI, false);
    sheet.getRange(sheet.getLastRow() + 1, 1, output.length, HEADER_TRANSFER_PRODUKSI.length).setValues(output);

    appendAuditLog_(ss, requestData, 'SUCCESS', 'Transfer produk outlet disimpan. ID: ' + txId + '. Baris: ' + output.length + '.');
    return dataResponse_({
      success: true,
      message: 'Transfer produk outlet berhasil disimpan dengan ID: ' + txId
    });
  });
}


/**
 * v14.3: data surat jalan dari transaksi transfer yang SUDAH disimpan.
 * Filter wajib: tanggal + outlet. Bisa pilih per ID Transfer atau gabungan semua transfer tanggal/outlet itu.
 */
function handleGetTransferPrintData_(ss, requestData) {
  var profile = requestData._user;
  if (profile && !profileHasPermission_(profile, 'transferProduksi') && !profileHasPermission_(profile, 'produksi') && !profileHasPermission_(profile, 'produkOutlet')) {
    return dataResponse_({ success: false, message: 'Akses cetak transfer ditolak untuk user login.' });
  }

  var tanggal = asText_(requestData.tanggal || requestData.date);
  var outlet = asText_(requestData.outlet);
  var txFilter = asText_(requestData.txId || requestData.idTransfer || '');

  if (!tanggal) return dataResponse_({ success: false, message: 'Tanggal print wajib diisi.' });
  if (!outlet) return dataResponse_({ success: false, message: 'Outlet print wajib dipilih.' });

  var sheet = ss.getSheetByName(CONFIG.SHEETS.TRANSFER_PRODUKSI);
  if (!sheet || sheet.getLastRow() < 2) {
    return dataResponse_({ success: true, groups: [], total: 0, message: 'Belum ada transfer yang tersimpan.' });
  }

  var data = sheet.getDataRange().getValues();
  var groupsMap = {};
  var order = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowTx = asText_(row[0]);
    var rowTanggal = formatSheetDate_(row[1]);
    var rowPetugas = asText_(row[2]);
    var rowOutlet = asText_(row[3]);
    var rowPic = asText_(row[4]);
    var idProduk = asText_(row[5]);
    var namaProduk = asText_(row[6]);
    var qty = parseNumber_(row[7]);
    var satuan = asText_(row[8]) || '-';
    var catatan = asText_(row[9]);

    if (rowTanggal !== tanggal) continue;
    if (normalizeKey_(rowOutlet) !== normalizeKey_(outlet)) continue;
    if (txFilter && normalizeKey_(txFilter) !== 'all' && rowTx !== txFilter) continue;
    if (qty <= 0) continue;

    if (!groupsMap[rowTx]) {
      groupsMap[rowTx] = {
        txId: rowTx,
        tanggal: rowTanggal,
        outlet: rowOutlet,
        pic: rowPic,
        petugas: rowPetugas,
        totalQty: 0,
        totalItem: 0,
        catatan: '',
        rows: []
      };
      order.push(rowTx);
    }

    groupsMap[rowTx].totalQty += qty;
    groupsMap[rowTx].totalItem += 1;
    if (catatan && groupsMap[rowTx].catatan.indexOf(catatan) === -1) {
      groupsMap[rowTx].catatan += (groupsMap[rowTx].catatan ? '; ' : '') + catatan;
    }
    groupsMap[rowTx].rows.push({
      idProduk: idProduk,
      namaProduk: namaProduk,
      qtyTransfer: qty,
      satuan: satuan,
      catatan: catatan
    });
  }

  var groups = order.map(function(tx) { return groupsMap[tx]; });
  groups.sort(function(a, b) { return String(b.txId).localeCompare(String(a.txId)); });

  // v14.9: data tambahan untuk format A4 Surat Jalan.
  // A4 harus menampilkan SEMUA produk aktif outlet, bukan hanya produk yang dikirim.
  // Stok Akhir = stok outlet sebelum dokumen transfer ini.
  // Tambah Produk = qty transfer pada dokumen ini.
  // Total Produk = Stok Akhir + Tambah Produk.
  groups.forEach(function(g) {
    g.a4Rows = buildTransferSuratJalanA4Rows_(ss, outlet, tanggal, g.rows || [], g.txId);
  });

  var allGroup = null;
  if (groups.length > 1) {
    var allRows = [];
    var allQty = 0;
    var allCatatan = '';
    groups.forEach(function(g) {
      (g.rows || []).forEach(function(r) { allRows.push(r); });
      allQty += parseNumber_(g.totalQty || 0);
      if (g.catatan && allCatatan.indexOf(g.catatan) === -1) {
        allCatatan += (allCatatan ? '; ' : '') + g.catatan;
      }
    });
    allGroup = {
      txId: 'ALL',
      tanggal: tanggal,
      outlet: outlet,
      pic: 'Gabungan',
      petugas: 'Gabungan',
      totalQty: allQty,
      totalItem: allRows.length,
      catatan: allCatatan,
      rows: allRows,
      a4Rows: buildTransferSuratJalanA4Rows_(ss, outlet, tanggal, allRows, 'ALL')
    };
  }

  return dataResponse_({ success: true, groups: groups, allGroup: allGroup, total: groups.length });
}

function buildTransferSuratJalanA4Rows_(ss, outlet, tanggal, transferRows, txId) {
  var products = getMasterProdukOutlet_(ss, outlet);
  var moves = buildProdukOutletMovementMaps_(ss, outlet, tanggal);
  var tambahMap = {};
  var transferredProductMap = {};

  (transferRows || []).forEach(function(r) {
    var key = productKey_(r.idProduk, r.namaProduk);
    if (!key) return;
    addMapQty_(tambahMap, key, r.qtyTransfer);
    if (!transferredProductMap[key]) {
      transferredProductMap[key] = {
        id: asText_(r.idProduk),
        nama: asText_(r.namaProduk),
        satuan: asText_(r.satuan) || '-',
        urutan: 99999
      };
    }
  });

  // Kalau ada data transfer lama yang produknya belum tercatat di MASTER_PRODUK_OUTLET,
  // tetap tampilkan di A4 agar surat jalan tidak kehilangan baris.
  var exists = {};
  products.forEach(function(p) { exists[productKey_(p.id, p.nama)] = true; });
  Object.keys(transferredProductMap).forEach(function(key) {
    if (!exists[key]) products.push(transferredProductMap[key]);
  });

  var beforeTxMap = (txId && txId !== 'ALL') ? buildTransferTodayBeforeTxMap_(ss, outlet, tanggal, txId) : {};

  products.sort(function(a, b) {
    var au = Number(a.urutan || 9999);
    var bu = Number(b.urutan || 9999);
    if (au !== bu) return au - bu;
    return (a.nama || '').localeCompare(b.nama || '');
  });

  return products.map(function(p) {
    var key = productKey_(p.id, p.nama);
    var stokAwalHari = (moves.transferBefore[key] || 0) - (moves.keluarBefore[key] || 0) + (moves.opnameBefore[key] || 0);
    var stokSebelumDokumen = stokAwalHari + (beforeTxMap[key] || 0);
    var tambah = tambahMap[key] || 0;
    return {
      idProduk: p.id,
      namaProduk: p.nama,
      satuan: p.satuan || '-',
      stokAkhir: stokSebelumDokumen,
      tambahProduk: tambah,
      totalProduk: stokSebelumDokumen + tambah,
      produkTerjual: '',
      produkAkhir: ''
    };
  });
}

function buildTransferTodayBeforeTxMap_(ss, outlet, tanggal, txId) {
  var map = {};
  var names = [CONFIG.SHEETS.TRANSFER_PRODUKSI, 'TRANSFER_PRODUKSI_STOK'];
  var seen = {};
  for (var n = 0; n < names.length; n++) {
    var sheetName = names[n];
    if (seen[sheetName]) continue;
    seen[sheetName] = true;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) continue;
    var data = sheet.getDataRange().getValues();
    var hm = buildHeaderMap_(data[0]);
    var ixTx = findHeaderIndex_(hm, ['idtransfer', 'idtransaksi', 'id'], 0);
    var ixTanggal = findHeaderIndex_(hm, ['tanggal'], 1);
    var ixOutlet = findHeaderIndex_(hm, ['outlet', 'outlettujuan'], 3);
    var ixId = findHeaderIndex_(hm, ['idproduk', 'idbarang', 'id'], 5);
    var ixNama = findHeaderIndex_(hm, ['namaproduk', 'produk', 'namabarang', 'nama'], 6);
    var ixQty = findHeaderIndex_(hm, ['qtytransfer', 'qtymasuk', 'qty'], 7);
    for (var i = 1; i < data.length; i++) {
      var rowTx = asText_(data[i][ixTx]);
      if (!rowTx || String(rowTx) >= String(txId)) continue;
      if (formatSheetDate_(data[i][ixTanggal]) !== tanggal) continue;
      if (normalizeKey_(data[i][ixOutlet]) !== normalizeKey_(outlet)) continue;
      var key = productKey_(data[i][ixId], data[i][ixNama]);
      if (!key) continue;
      addMapQty_(map, key, data[i][ixQty]);
    }
  }
  return map;
}

function hasTransferProduksiAccessByName_(ss, userName) {
  var userSheet = ss.getSheetByName(CONFIG.SHEETS.USER);
  if (!userSheet) return false;

  var users = userSheet.getDataRange().getValues();
  var level = '';
  for (var i = 1; i < users.length; i++) {
    if (normalizeKey_(users[i][3]) === normalizeKey_(userName)) {
      level = asText_(users[i][4]);
      break;
    }
  }
  if (!level) return false;
  if (isOwnerRole_(level)) return true;

  var p = readPermissionsForLevel_(ss, level);
  // Saat kolom TransferProduksi belum dibuat, Produksi=Y boleh transfer agar DAPUR tidak terblokir.
  return permissionYes_(p, 'transferProduksi') || permissionYes_(p, 'produksi');
}

/**
 * REVISI v6: Cek permission langsung dari sheet berdasarkan nama petugas.
 * Dipakai untuk proteksi backend, supaya user tidak bisa simpan Produksi hanya dari manipulasi frontend/localStorage.
 */
function hasUserPermissionByName_(ss, userName, permissionKey) {
  var userSheet = ss.getSheetByName(CONFIG.SHEETS.USER);
  if (!userSheet) return false;

  var users = userSheet.getDataRange().getValues();
  var level = '';
  for (var i = 1; i < users.length; i++) {
    if (normalizeKey_(users[i][3]) === normalizeKey_(userName)) {
      level = asText_(users[i][4]);
      break;
    }
  }
  if (!level) return false;

  var permissionMap = readPermissionsForLevel_(ss, level);
  return permissionYes_(permissionMap, permissionKey);
}

/**
 * DASHBOARD v4 - Ringkasan berbasis role.
 * Owner melihat ikhtisar seluruh gudang.
 * Selain Owner hanya melihat aktivitas dirinya + menu sesuai permission.
 */
function handleGetDashboardData_(ss, requestData) {
  var profile = requestData._user || {};
  if (!profile.nama) {
    return dataResponse_({
      success: false,
      code: 'SESSION_EXPIRED',
      message: 'Sesi tidak valid atau user tidak terbaca. Silakan login ulang.'
    });
  }
  var userName = asText_(profile.nama);
  var level = asText_(profile.level);
  var permissions = profile.permissions || readPermissionsForLevel_(ss, level);
  var isOwner = isOwnerRole_(level);
  var isAdmin = isAdminRole_(level);

  var today = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd');
  var monthKey = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM');

  var allStockItems = readStockItems_(ss);
  var visibleStockItems = isOwner ? allStockItems : filterStockByPermission_(allStockItems, permissions);

  var stockHealth = summarizeStockHealthV9_(visibleStockItems);
  var txSummary = summarizeTransactions_(ss, today, monthKey, userName, !isOwner);
  var categorySummary = summarizeCategories_(visibleStockItems);
  var locationSummary = summarizeLocations_(visibleStockItems);
  var criticalItems = getCriticalStockItems_(visibleStockItems, isOwner ? 14 : 10);
  var recentActivities = getRecentActivities_(ss, userName, !isOwner, isOwner ? 14 : 10);
  var topMovements = getTopTodayMovements_(ss, today, userName, !isOwner, 8);
  var quickMenus = buildQuickMenus_(isOwner, permissions, isAdmin);
  var insights = buildDashboardInsights_(stockHealth, txSummary, criticalItems, locationSummary, isOwner);

  return dataResponse_({
    success: true,
    version: 'v9-owner-control-panel',
    mode: isOwner ? 'owner' : 'operator',
    today: today,
    month: monthKey,
    userName: userName,
    level: level,
    quickMenus: quickMenus,
    kpi: {
      transaksiHariIni: txSummary.totalToday,
      inputHariIni: txSummary.inputToday,
      outputHariIni: txSummary.outputToday,
      produksiHariIni: txSummary.produksiToday,
      transferHariIni: txSummary.transferToday,
      opnameBulanIni: txSummary.opnameMonth,
      totalItem: stockHealth.totalItem,
      stokAman: stockHealth.aman,
      stokWaspada: stockHealth.waspada,
      stokKritis: stockHealth.kritis,
      stokKosong: stockHealth.kosong,
      lokasiAktif: locationSummary.length,
      itemPusat: stockHealth.itemPusat,
      itemOutlet: stockHealth.itemOutlet,
      itemProduksi: stockHealth.itemProduksi,
      aksesAktif: quickMenus.filter(function(menu) { return !menu.disabled; }).length
    },
    categorySummary: categorySummary,
    locationSummary: locationSummary,
    criticalItems: criticalItems,
    recentActivities: recentActivities,
    topMovements: topMovements,
    insights: insights
  });
}

function isOwnerRole_(level) {
  var lv = normalizeKey_(level);
  return lv === 'owner' || lv === 'superadmin' || lv === 'super admin';
}

function isAdminRole_(level) {
  var lv = normalizeKey_(level);
  return isOwnerRole_(level) || lv === 'supervisor';
}

function permissionYes_(permissions, key) {
  return asYN_(permissions && permissions[key]) === 'Y';
}

function categoryPermissionKey_(kategori) {
  var k = normalizeKey_(kategori);
  if (k.indexOf('mentah') !== -1) return 'bahanMentah';
  if (k.indexOf('baku') !== -1) return 'bahanBaku';
  if (k.indexOf('kemasan') !== -1) return 'kemasan';
  if (k.indexOf('minuman') !== -1) return 'minuman';
  if (k.indexOf('bersih') !== -1) return 'kebersihan';
  if (k.indexOf('produksi') !== -1) return 'produksi';
  return '';
}

function isCategoryAllowed_(kategori, permissions) {
  var key = categoryPermissionKey_(kategori);
  return key ? permissionYes_(permissions, key) : false;
}

function filterStockByPermission_(items, permissions) {
  // REVISI v7:
  // Lihat Stok adalah izin baca global. Jadi role seperti LEVEL_1 yang semua kategori N
  // tetapi Lihat Stok = Y tetap bisa membuka dan melihat laporan stok.
  // Untuk membatasi lihat stok per kategori, nanti tambahkan kolom izin khusus per kategori view.
  if (permissionYes_(permissions, 'lihatStok')) return items.slice();

  var result = [];
  for (var i = 0; i < items.length; i++) {
    if (isCategoryAllowed_(items[i].kategori, permissions)) result.push(items[i]);
  }
  return result;
}

function buildQuickMenus_(isOwner, permissions, isAdmin) {
  var menus = [];
  menus.push({ label: 'Dashboard', href: 'dashboard.html', tone: 'blue', desc: 'Ikhtisar sistem' });

  var canGeneralStock = isOwner || permissionYes_(permissions, 'bahanMentah') || permissionYes_(permissions, 'bahanBaku') || permissionYes_(permissions, 'kemasan') || permissionYes_(permissions, 'minuman') || permissionYes_(permissions, 'kebersihan');
  var canOutput = canGeneralStock || permissionYes_(permissions, 'produksi');
  var canProduksi = isOwner || permissionYes_(permissions, 'produksi');
  var canTransferProduksi = isOwner || permissionYes_(permissions, 'transferProduksi') || permissionYes_(permissions, 'produksi');
  var canProdukOutlet = isOwner || permissionYes_(permissions, 'produkOutlet');

  if (canGeneralStock) menus.push({ label: 'Input Stok', href: 'input-stok.html', tone: 'blue', desc: 'Barang masuk pusat' });
  if (canOutput) menus.push({ label: 'Output Stok', href: 'output-stok.html', tone: 'rose', desc: 'Barang keluar / produksi' });
  if (canTransferProduksi) menus.push({ label: 'Transfer Produk Outlet', href: 'transfer-produksi.html', tone: 'violet', desc: 'Kirim produk dari PUSAT ke outlet' });
  if (canProdukOutlet) menus.push({ label: 'Produk Outlet', href: 'produk-outlet.html', tone: 'sky', desc: 'Terjual & opname outlet' });
  if (isOwner || permissionYes_(permissions, 'opname')) menus.push({ label: 'Stok Opname', href: 'stok-opname.html', tone: 'amber', desc: 'Koreksi stok fisik' });
  if (isOwner || permissionYes_(permissions, 'lihatStok')) menus.push({ label: 'Lihat Stok', href: 'lihat-stok.html', tone: 'emerald', desc: 'Laporan stok per lokasi' });
  if (isOwner || permissionYes_(permissions, 'lihatStok')) menus.push({ label: 'Riwayat Transaksi', href: 'riwayat-transaksi.html', tone: 'blue', desc: 'Audit transaksi stok' });
  if (isAdmin || isOwner) menus.push({ label: 'Admin Sistem', href: 'admin.html', tone: 'violet', desc: 'User, password, permission, dan audit' });
  if (isOwner) menus.push({ label: 'Arsip Data', href: '#', tone: 'slate', desc: 'Segera: backup rentang tanggal', disabled: true });

  return menus;
}

function requireAdminSession_(ss, requestData) {
  var profile = getSessionProfile_(ss, requestData);
  if (!profile || !profile.nama || !isAdminRole_(profile.level)) return null;
  requestData._user = profile;
  return profile;
}

function adminDeniedResponse_() {
  return dataResponse_({
    success: false,
    code: 'ADMIN_DENIED',
    message: 'Akses admin ditolak. Silakan login memakai akun Owner/SuperAdmin/Supervisor.'
  });
}

function handleGetAdminData_(ss, requestData) {
  var admin = requireAdminSession_(ss, requestData);
  if (!admin) return adminDeniedResponse_();

  return dataResponse_({
    success: true,
    admin: {
      username: admin.username,
      nama: admin.nama,
      level: admin.level,
      outlet: admin.outlet
    },
    users: readAdminUsers_(ss),
    permissionColumns: getAdminPermissionColumns_(),
    permissionRows: readAdminPermissionRows_(ss),
    auditLogs: readAuditLogs_(ss, 80),
    masterSummary: readAdminMasterSummary_(ss)
  });
}

function handleAdminUpdatePassword_(ss, requestData) {
  var admin = requireAdminSession_(ss, requestData);
  if (!admin) return adminDeniedResponse_();

  var username = asText_(requestData.username);
  var newPassword = asText_(requestData.newPassword);
  if (!username) return dataResponse_({ success: false, message: 'Username wajib dipilih.' });
  if (newPassword.length < 4) return dataResponse_({ success: false, message: 'Password baru minimal 4 karakter.' });

  return withScriptLock_(function() {
    var sheet = ss.getSheetByName(CONFIG.SHEETS.USER);
    if (!sheet || sheet.getLastRow() < 2) return dataResponse_({ success: false, message: 'Tabel USER tidak ditemukan atau masih kosong.' });

    var values = sheet.getDataRange().getDisplayValues();
    var headerMap = buildHeaderMap_(values[0]);
    var ixUsername = findHeaderIndex_(headerMap, ['username', 'user', 'namauser'], 1);
    var ixPassword = findHeaderIndex_(headerMap, ['password', 'pass', 'sandi'], 2);
    var targetRow = -1;

    for (var r = 1; r < values.length; r++) {
      if (normalizeKey_(values[r][ixUsername]) === normalizeKey_(username)) {
        targetRow = r + 1;
        break;
      }
    }

    if (targetRow < 0) return dataResponse_({ success: false, message: 'Username tidak ditemukan di tabel USER.' });

    sheet.getRange(targetRow, ixPassword + 1).setValue(buatHashPasswordDenganSalt(newPassword));
    appendAuditLog_(ss, { action: 'adminUpdatePassword', _user: admin }, 'SUCCESS', 'Password user ' + username + ' diubah ke hash.');
    return dataResponse_({ success: true, message: 'Password user ' + username + ' berhasil diubah ke hash.' });
  });
}

function handleAdminSavePermissions_(ss, requestData) {
  var admin = requireAdminSession_(ss, requestData);
  if (!admin) return adminDeniedResponse_();

  var role = asText_(requestData.role);
  var permissions = requestData.permissions || {};
  if (!role) return dataResponse_({ success: false, message: 'Role wajib dipilih.' });
  if (isOwnerRole_(role)) return dataResponse_({ success: false, message: 'Role Owner/SuperAdmin selalu mendapat akses penuh otomatis.' });

  return withScriptLock_(function() {
    var sheet = ss.getSheetByName(CONFIG.SHEETS.PERMISSION);
    if (!sheet || sheet.getLastRow() < 2) return dataResponse_({ success: false, message: 'Tabel LEVEL_PERMISSION tidak ditemukan atau masih kosong.' });

    var values = sheet.getDataRange().getDisplayValues();
    var headers = values[0] || [];
    var colByKey = {};
    for (var c = 0; c < headers.length; c++) {
      var key = permissionKeyFromHeader_(headers[c]);
      if (key) colByKey[key] = c;
    }

    var targetRowIndex = -1;
    for (var r = 1; r < values.length; r++) {
      if (normalizeKey_(values[r][0]) === normalizeKey_(role)) {
        targetRowIndex = r;
        break;
      }
    }

    if (targetRowIndex < 0) return dataResponse_({ success: false, message: 'Role tidak ditemukan di LEVEL_PERMISSION.' });

    var row = values[targetRowIndex].slice();
    var updated = 0;
    getAdminPermissionColumns_().forEach(function(col) {
      if (!Object.prototype.hasOwnProperty.call(permissions, col.key)) return;
      if (typeof colByKey[col.key] === 'undefined') return;
      row[colByKey[col.key]] = asYN_(permissions[col.key]);
      updated++;
    });

    if (!updated) return dataResponse_({ success: false, message: 'Tidak ada kolom permission yang cocok untuk disimpan.' });

    sheet.getRange(targetRowIndex + 1, 1, 1, headers.length).setValues([row.slice(0, headers.length)]);
    appendAuditLog_(ss, { action: 'adminSavePermissions', _user: admin }, 'SUCCESS', 'Permission role ' + role + ' diperbarui.');
    return dataResponse_({ success: true, message: 'Permission role ' + role + ' berhasil disimpan.' });
  });
}

function readAdminUsers_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.USER);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getDataRange().getDisplayValues();
  var headerMap = buildHeaderMap_(data[0]);
  var ixUsername = findHeaderIndex_(headerMap, ['username', 'user', 'namauser'], 1);
  var ixPassword = findHeaderIndex_(headerMap, ['password', 'pass', 'sandi'], 2);
  var ixNama = findHeaderIndex_(headerMap, ['nama', 'namakaryawan', 'petugas'], 3);
  var ixLevel = findHeaderIndex_(headerMap, ['level', 'role', 'akses'], 4);
  var ixOutlet = findHeaderIndex_(headerMap, ['outlet', 'lokasi', 'cabang'], 5);
  var users = [];

  for (var r = 1; r < data.length; r++) {
    var username = asText_(data[r][ixUsername]);
    if (!username) continue;
    var password = asText_(data[r][ixPassword]);
    users.push({
      rowNumber: r + 1,
      username: username,
      nama: asText_(data[r][ixNama]),
      level: asText_(data[r][ixLevel]),
      outlet: asText_(data[r][ixOutlet]),
      passwordMode: password ? (/^sha256\$/i.test(password) ? 'HASH' : 'PLAIN') : 'EMPTY'
    });
  }
  return users;
}

function getAdminPermissionColumns_() {
  return [
    { key: 'bahanMentah', label: 'Bahan Mentah' },
    { key: 'bahanBaku', label: 'Bahan Baku' },
    { key: 'kemasan', label: 'Kemasan' },
    { key: 'minuman', label: 'Minuman' },
    { key: 'kebersihan', label: 'Kebersihan' },
    { key: 'opname', label: 'Opname' },
    { key: 'produksi', label: 'Produksi' },
    { key: 'transferProduksi', label: 'Transfer Produksi' },
    { key: 'lihatStok', label: 'Lihat Stok' },
    { key: 'produkOutlet', label: 'Produk Outlet' }
  ];
}

function readAdminPermissionRows_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PERMISSION);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getDataRange().getDisplayValues();
  var headers = data[0] || [];
  var colByKey = {};
  for (var c = 0; c < headers.length; c++) {
    var key = permissionKeyFromHeader_(headers[c]);
    if (key) colByKey[key] = c;
  }

  var rows = [];
  for (var r = 1; r < data.length; r++) {
    var role = asText_(data[r][0]);
    if (!role) continue;
    var perms = {};
    getAdminPermissionColumns_().forEach(function(col) {
      perms[col.key] = typeof colByKey[col.key] !== 'undefined'
        ? asYN_(data[r][colByKey[col.key]])
        : '';
    });
    rows.push({
      rowNumber: r + 1,
      role: role,
      locked: isOwnerRole_(role),
      permissions: perms
    });
  }
  return rows;
}

function readAuditLogs_(ss, limit) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.AUDIT_LOG);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var count = Math.min(Number(limit) || 50, sheet.getLastRow() - 1);
  var startRow = sheet.getLastRow() - count + 1;
  var data = sheet.getRange(startRow, 1, count, HEADER_AUDIT_LOG.length).getDisplayValues();
  var logs = data.map(function(row) {
    return {
      timestamp: asText_(row[0]),
      action: asText_(row[1]),
      status: asText_(row[2]),
      user: asText_(row[3]),
      level: asText_(row[4]),
      outlet: asText_(row[5]),
      message: asText_(row[6]),
      requestId: asText_(row[7])
    };
  });
  logs.reverse();
  return logs;
}

function readAdminMasterSummary_(ss) {
  var sheetNames = [
    'MASTER_BAHAN_MENTAH',
    'MASTER_BAHAN_BAKU',
    'MASTER_KEMASAN',
    'MASTER_MINUMAN',
    'MASTER_KEBERSIHAN',
    CONFIG.SHEETS.MASTER_PRODUKSI,
    CONFIG.SHEETS.MASTER_PRODUK_OUTLET,
    CONFIG.SHEETS.STOK_AKHIR,
    CONFIG.SHEETS.OUTLET,
    CONFIG.SHEETS.USER,
    CONFIG.SHEETS.PERMISSION
  ];

  return sheetNames.map(function(name) {
    var sheet = ss.getSheetByName(name);
    return {
      name: name,
      exists: !!sheet,
      rows: sheet ? Math.max(0, sheet.getLastRow() - 1) : 0,
      columns: sheet ? sheet.getLastColumn() : 0
    };
  });
}

function readStockItems_(ss) {
  var maps = buildStockMaps_(ss);
  var hiddenParentIds = getBahanMentahParentIdMap_(ss);
  var result = [];

  for (var i = 0; i < maps.allRows.length; i++) {
    var item = maps.allRows[i];
    if (!item.id && !item.nama) continue;
    if (item.id && hiddenParentIds[normalizeKey_(item.id)]) continue;

    result.push({
      lokasi: item.lokasi || CONFIG.DEFAULT_LOCATION,
      kategori: displayCategoryLabel_(item.kategori || 'Tanpa Kategori'),
      id: item.id,
      nama: item.nama,
      satuan: item.satuan || '-',
      qty: parseNumber_(item.qty),
      qtyDisplay: asText_(item.qty) || '0',
      lastUpdate: item.lastUpdate,
      status: item.status
    });
  }
  return result;
}

function getBahanMentahParentIdMap_(ss) {
  var map = {};
  var sheet = ss.getSheetByName('MASTER_BAHAN_MENTAH');
  if (!sheet || sheet.getLastRow() < 2) return map;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var id = asText_(data[i][0]);
    var tipe = asText_(data[i][5]).toUpperCase();
    if (id && tipe === 'PARENT') map[normalizeKey_(id)] = true;
  }
  return map;
}

function stockLevel_(item) {
  var status = normalizeKey_(item.status);
  var qty = parseNumber_(item.qty);

  if (qty <= 0 || status.indexOf('habis') !== -1 || status.indexOf('kosong') !== -1 || status.indexOf('kritis') !== -1) return 'kritis';
  if (status.indexOf('minimum') !== -1 || status.indexOf('menipis') !== -1 || status.indexOf('low') !== -1 || status.indexOf('kurang') !== -1 || status.indexOf('waspada') !== -1) return 'waspada';
  return 'aman';
}

function summarizeStockHealth_(items) {
  var res = { totalItem: items.length, aman: 0, waspada: 0, kritis: 0 };
  for (var i = 0; i < items.length; i++) {
    var level = stockLevel_(items[i]);
    if (level === 'kritis') res.kritis++;
    else if (level === 'waspada') res.waspada++;
    else res.aman++;
  }
  return res;
}

function summarizeCategories_(items) {
  var grouped = {};
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var key = item.kategori || 'Tanpa Kategori';
    if (!grouped[key]) grouped[key] = { kategori: key, totalItem: 0, aman: 0, waspada: 0, kritis: 0 };
    grouped[key].totalItem++;
    var level = stockLevel_(item);
    grouped[key][level]++;
  }
  var out = Object.keys(grouped).map(function(key) { return grouped[key]; });
  out.sort(function(a, b) { return a.kategori.localeCompare(b.kategori); });
  return out;
}


function summarizeStockHealthV9_(items) {
  var res = {
    totalItem: items.length,
    aman: 0,
    waspada: 0,
    kritis: 0,
    kosong: 0,
    itemPusat: 0,
    itemOutlet: 0,
    itemProduksi: 0
  };
  var defaultLoc = normalizeKey_(CONFIG.DEFAULT_LOCATION);
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var qty = parseNumber_(item.qty);
    var level = stockLevel_(item);
    if (qty <= 0) res.kosong++;
    if (level === 'kritis') res.kritis++;
    else if (level === 'waspada') res.waspada++;
    else res.aman++;
    if (normalizeKey_(item.lokasi) === defaultLoc) res.itemPusat++;
    else res.itemOutlet++;
    if (normalizeKey_(item.kategori).indexOf('produksi') !== -1) res.itemProduksi++;
  }
  return res;
}

function summarizeLocations_(items) {
  var grouped = {};
  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    var key = item.lokasi || CONFIG.DEFAULT_LOCATION;
    if (!grouped[key]) grouped[key] = { lokasi: key, totalItem: 0, aman: 0, waspada: 0, kritis: 0, kosong: 0, produksi: 0, qtyProduksi: 0 };
    grouped[key].totalItem++;
    var level = stockLevel_(item);
    grouped[key][level]++;
    if (parseNumber_(item.qty) <= 0) grouped[key].kosong++;
    if (normalizeKey_(item.kategori).indexOf('produksi') !== -1) {
      grouped[key].produksi++;
      grouped[key].qtyProduksi += parseNumber_(item.qty);
    }
  }
  var out = Object.keys(grouped).map(function(key) { return grouped[key]; });
  out.sort(function(a, b) {
    var aPusat = normalizeKey_(a.lokasi) === normalizeKey_(CONFIG.DEFAULT_LOCATION) ? 0 : 1;
    var bPusat = normalizeKey_(b.lokasi) === normalizeKey_(CONFIG.DEFAULT_LOCATION) ? 0 : 1;
    if (aPusat !== bPusat) return aPusat - bPusat;
    if (a.kritis !== b.kritis) return b.kritis - a.kritis;
    if (a.waspada !== b.waspada) return b.waspada - a.waspada;
    return a.lokasi.localeCompare(b.lokasi);
  });
  return out;
}

function buildDashboardInsights_(health, tx, criticalItems, locations, isOwner) {
  var insights = [];
  if (health.kosong > 0) {
    insights.push({ tone: 'rose', title: 'Ada stok kosong', text: health.kosong + ' item stok tercatat kosong. Prioritaskan cek fisik dan rencana belanja/produksi.' });
  } else if (health.kritis > 0) {
    insights.push({ tone: 'rose', title: 'Ada stok kritis', text: health.kritis + ' item masuk status kritis. Cek daftar prioritas di panel stok perhatian.' });
  } else if (health.waspada > 0) {
    insights.push({ tone: 'amber', title: 'Stok waspada', text: health.waspada + ' item mulai menipis. Masih aman, tapi perlu disiapkan sebelum ramai.' });
  } else {
    insights.push({ tone: 'emerald', title: 'Stok aman', text: 'Belum ada stok kritis atau waspada pada akses dashboard ini.' });
  }

  if (tx.totalToday === 0) {
    insights.push({ tone: 'slate', title: 'Belum ada transaksi hari ini', text: isOwner ? 'Dashboard belum menerima input/output/produksi/transfer hari ini.' : 'Belum ada transaksi atas nama login ini hari ini.' });
  } else {
    insights.push({ tone: 'blue', title: 'Aktivitas hari ini berjalan', text: tx.totalToday + ' baris transaksi tercatat hari ini. Input ' + tx.inputToday + ', output ' + tx.outputToday + ', produksi ' + tx.produksiToday + ', transfer ' + tx.transferToday + '.' });
  }

  if (locations && locations.length) {
    var worst = locations.slice().sort(function(a, b) {
      if (a.kritis !== b.kritis) return b.kritis - a.kritis;
      if (a.waspada !== b.waspada) return b.waspada - a.waspada;
      return b.totalItem - a.totalItem;
    })[0];
    insights.push({ tone: worst.kritis > 0 ? 'rose' : (worst.waspada > 0 ? 'amber' : 'emerald'), title: 'Lokasi prioritas: ' + worst.lokasi, text: worst.totalItem + ' item, ' + worst.kritis + ' kritis, ' + worst.waspada + ' waspada, ' + worst.kosong + ' kosong.' });
  }

  return insights.slice(0, 4);
}

function getTopTodayMovements_(ss, today, userName, userOnly, limit) {
  var map = {};
  addMovementFromSheet_(ss, map, 'TRANSAKSI_INPUT_STOK', 'Input', today, userName, userOnly, {
    tanggal: 1, petugas: 2, nama: 5, qty: 8, satuan: 9, lokasi: null
  });
  addMovementFromSheet_(ss, map, CONFIG.SHEETS.OUTPUT_STOK, 'Output', today, userName, userOnly, {
    tanggal: 1, petugas: 2, nama: 5, qty: 6, satuan: 7, lokasi: 8
  });
  addMovementFromSheet_(ss, map, CONFIG.SHEETS.OUTPUT_PRODUKSI, 'Produksi', today, userName, userOnly, {
    tanggal: 1, petugas: 2, nama: 4, qty: 5, satuan: 6, lokasi: null
  });
  addMovementFromSheet_(ss, map, CONFIG.SHEETS.TRANSFER_PRODUKSI, 'Transfer', today, userName, userOnly, {
    tanggal: 1, petugas: 2, nama: 6, qty: 7, satuan: 8, lokasi: 3
  });

  var out = Object.keys(map).map(function(key) { return map[key]; });
  out.sort(function(a, b) { return parseNumber_(b.qty) - parseNumber_(a.qty); });
  return out.slice(0, limit || 8);
}

function addMovementFromSheet_(ss, targetMap, sheetName, jenis, today, userName, userOnly, map) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var take = Math.min(500, lastRow - 1);
  var startRow = lastRow - take + 1;
  var data = sheet.getRange(startRow, 1, take, lastCol).getValues();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (formatSheetDate_(row[map.tanggal]) !== today) continue;
    var petugas = asText_(row[map.petugas]);
    if (userOnly && normalizeKey_(petugas) !== normalizeKey_(userName)) continue;
    var nama = asText_(row[map.nama]);
    var satuan = asText_(row[map.satuan]);
    var lokasi = map.lokasi === null || typeof map.lokasi === 'undefined' ? CONFIG.DEFAULT_LOCATION : asText_(row[map.lokasi]);
    var qty = parseNumber_(row[map.qty]);
    if (!nama || !qty) continue;
    var key = jenis + '|' + normalizeKey_(nama) + '|' + normalizeKey_(satuan) + '|' + normalizeKey_(lokasi);
    if (!targetMap[key]) targetMap[key] = { jenis: jenis, nama: nama, qty: 0, satuan: satuan, lokasi: lokasi || CONFIG.DEFAULT_LOCATION };
    targetMap[key].qty += qty;
  }
}

function getCriticalStockItems_(items, limit) {
  var list = [];
  for (var i = 0; i < items.length; i++) {
    var level = stockLevel_(items[i]);
    if (level === 'kritis' || level === 'waspada') {
      list.push({
        lokasi: items[i].lokasi || CONFIG.DEFAULT_LOCATION,
        kategori: items[i].kategori,
        id: items[i].id,
        nama: items[i].nama,
        qty: items[i].qtyDisplay || String(items[i].qty),
        satuan: items[i].satuan,
        status: items[i].status || (level === 'kritis' ? 'Kritis' : 'Waspada'),
        level: level
      });
    }
  }
  list.sort(function(a, b) {
    if (a.level !== b.level) return a.level === 'kritis' ? -1 : 1;
    return parseNumber_(a.qty) - parseNumber_(b.qty);
  });
  return list.slice(0, limit || 10);
}

function summarizeTransactions_(ss, today, monthKey, userName, userOnly) {
  var input = countSheetRows_(ss, 'TRANSAKSI_INPUT_STOK', today, 1, userName, 2, userOnly, false);
  var output = countSheetRows_(ss, CONFIG.SHEETS.OUTPUT_STOK, today, 1, userName, 2, userOnly, false);
  var produksi = countSheetRows_(ss, CONFIG.SHEETS.OUTPUT_PRODUKSI, today, 1, userName, 2, userOnly, false);
  var transfer = countSheetRows_(ss, CONFIG.SHEETS.TRANSFER_PRODUKSI, today, 1, userName, 2, userOnly, false);
  var opnameMonth = countSheetRows_(ss, CONFIG.SHEETS.STOK_OPNAME, monthKey, 1, userName, 2, userOnly, true);
  return {
    inputToday: input,
    outputToday: output,
    produksiToday: produksi,
    transferToday: transfer,
    opnameMonth: opnameMonth,
    totalToday: input + output + produksi + transfer
  };
}

function countSheetRows_(ss, sheetName, dateKey, dateIndex, userName, userIndex, userOnly, monthMode) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return 0;
  var data = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (userOnly && normalizeKey_(data[i][userIndex]) !== normalizeKey_(userName)) continue;
    var dateText = formatSheetDate_(data[i][dateIndex]);
    if (monthMode) {
      if (dateText.indexOf(dateKey) === 0) count++;
    } else {
      if (dateText === dateKey) count++;
    }
  }
  return count;
}

function getRecentActivities_(ss, userName, userOnly, limit) {
  var activities = [];
  addRecentFromSheet_(ss, activities, 'TRANSAKSI_INPUT_STOK', 'Input', userName, userOnly, {
    tx: 0, tanggal: 1, petugas: 2, kategori: 3, nama: 5, qty: 8, satuan: 9
  });
  addRecentFromSheet_(ss, activities, CONFIG.SHEETS.OUTPUT_STOK, 'Output', userName, userOnly, {
    tx: 0, tanggal: 1, petugas: 2, kategori: 3, nama: 5, qty: 6, satuan: 7, tujuan: 8
  });
  addRecentFromSheet_(ss, activities, CONFIG.SHEETS.OUTPUT_PRODUKSI, 'Produksi', userName, userOnly, {
    tx: 0, tanggal: 1, petugas: 2, kategori: null, nama: 4, qty: 5, satuan: 6
  });
  addRecentFromSheet_(ss, activities, CONFIG.SHEETS.TRANSFER_PRODUKSI, 'Transfer', userName, userOnly, {
    tx: 0, tanggal: 1, petugas: 2, kategori: null, nama: 6, qty: 7, satuan: 8, tujuan: 3
  });
  addRecentFromSheet_(ss, activities, CONFIG.SHEETS.STOK_OPNAME, 'Opname', userName, userOnly, {
    tx: 0, tanggal: 1, petugas: 2, kategori: 3, nama: 5, qty: 8, satuan: 9
  });
  addRecentFromSheet_(ss, activities, CONFIG.SHEETS.PRODUK_OUTLET_KELUAR, 'Produk Outlet Terjual', userName, userOnly, {
    tx: 0, tanggal: 1, petugas: 3, kategori: 2, nama: 5, qty: 6, satuan: 7, tujuan: 2
  });
  addRecentFromSheet_(ss, activities, CONFIG.SHEETS.OPNAME_PRODUK_OUTLET, 'Opname Produk Outlet', userName, userOnly, {
    tx: 0, tanggal: 1, petugas: 3, kategori: 2, nama: 5, qty: 8, satuan: 9, tujuan: 2
  });

  activities.sort(function(a, b) { return b.rank.localeCompare(a.rank); });
  var out = activities.slice(0, limit || 10);
  for (var i = 0; i < out.length; i++) delete out[i].rank;
  return out;
}

function addRecentFromSheet_(ss, target, sheetName, jenis, userName, userOnly, map) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var take = Math.min(300, lastRow - 1);
  var startRow = lastRow - take + 1;
  var data = sheet.getRange(startRow, 1, take, lastCol).getValues();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var petugas = asText_(row[map.petugas]);
    if (userOnly && normalizeKey_(petugas) !== normalizeKey_(userName)) continue;

    var txId = asText_(row[map.tx]);
    var tanggal = formatSheetDate_(row[map.tanggal]);
    target.push({
      jenis: jenis,
      txId: txId,
      tanggal: tanggal,
      petugas: petugas,
      kategori: map.kategori === null ? jenis : asText_(row[map.kategori]),
      nama: asText_(row[map.nama]),
      qty: asText_(row[map.qty]),
      satuan: asText_(row[map.satuan]),
      tujuan: map.tujuan === null || typeof map.tujuan === 'undefined' ? '' : asText_(row[map.tujuan]),
      rank: buildActivityRank_(tanggal, txId, i)
    });
  }
}

function buildActivityRank_(tanggal, txId, idx) {
  var cleanTx = asText_(txId).replace(/[^0-9]/g, '');
  return asText_(tanggal).replace(/[^0-9]/g, '') + cleanTx + ('000000' + idx).slice(-6);
}

/** Helper stok: membuat map byId dan byName dari STOK_AKHIR. */
function buildStockMaps_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.STOK_AKHIR);
  var maps = { byId: {}, byName: {}, byLocationId: {}, byLocationName: {}, allRows: [] };
  if (!sheet || sheet.getLastRow() < 2) return maps;

  var data = sheet.getDataRange().getValues();
  var headerMap = buildHeaderMap_(data[0]);
  var ix = getStokAkhirIndexes_(headerMap);

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var lokasi = asText_(row[ix.lokasi]) || CONFIG.DEFAULT_LOCATION;
    var kategori = asText_(row[ix.kategori]);
    var id = asText_(row[ix.id]);
    var nama = asText_(row[ix.nama]);
    var satuan = asText_(row[ix.satuan]);
    var qty = parseNumber_(row[ix.stokAkhir]);
    var lastUpdate = formatSheetDate_(row[ix.lastUpdate]);
    var status = asText_(row[ix.status]);

    if (!id && !nama) continue;

    var item = {
      lokasi: lokasi,
      kategori: kategori,
      id: id,
      nama: nama,
      satuan: satuan,
      qty: qty,
      lastUpdate: lastUpdate,
      status: status
    };

    maps.allRows.push(item);
    addStockItemToMaps_(maps, item);
  }
  return maps;
}

function getStokAkhirIndexes_(headerMap) {
  return {
    lokasi: findHeaderIndex_(headerMap, ['lokasi'], 0),
    kategori: findHeaderIndex_(headerMap, ['kategori'], 1),
    id: findHeaderIndex_(headerMap, ['idbarang', 'id'], 2),
    nama: findHeaderIndex_(headerMap, ['namabarang', 'nama'], 3),
    satuan: findHeaderIndex_(headerMap, ['satuan', 'satuanstok'], 4),
    stokAkhir: findHeaderIndex_(headerMap, ['stokakhir'], 13),
    lastUpdate: findHeaderIndex_(headerMap, ['lastupdate'], 14),
    status: findHeaderIndex_(headerMap, ['status'], 15)
  };
}

function addStockItemToMaps_(maps, item) {
  var locKey = normalizeKey_(item.lokasi || CONFIG.DEFAULT_LOCATION);
  var idKey = normalizeKey_(item.id);
  var nameKey = normalizeKey_(item.nama);
  var defaultLocKey = normalizeKey_(CONFIG.DEFAULT_LOCATION);

  if (idKey) maps.byLocationId[locKey + '|' + idKey] = item;
  if (nameKey) maps.byLocationName[locKey + '|' + nameKey] = item;

  // byId/byName diprioritaskan ke PUSAT agar Input/Output/Produksi tidak mengambil stok outlet.
  if (locKey === defaultLocKey) {
    if (idKey) maps.byId[idKey] = item;
    if (nameKey) maps.byName[nameKey] = item;
  } else {
    if (idKey && !maps.byId[idKey]) maps.byId[idKey] = item;
    if (nameKey && !maps.byName[nameKey]) maps.byName[nameKey] = item;
  }
}

function getStockItem_(maps, lokasi, id, nama) {
  var locKey = normalizeKey_(lokasi || CONFIG.DEFAULT_LOCATION);
  var idKey = normalizeKey_(id);
  var nameKey = normalizeKey_(nama);

  if (idKey && maps.byLocationId[locKey + '|' + idKey]) return maps.byLocationId[locKey + '|' + idKey];
  if (nameKey && maps.byLocationName[locKey + '|' + nameKey]) return maps.byLocationName[locKey + '|' + nameKey];
  if (idKey && maps.byId[idKey]) return maps.byId[idKey];
  if (nameKey && maps.byName[nameKey]) return maps.byName[nameKey];
  return null;
}

/**
 * BARU v3: fallback qty produksi dari TRANSAKSI_OUTPUT_PRODUKSI.
 * Ini bukan pengganti STOK_AKHIR, hanya cadangan agar laporan produksi tetap terbaca saat formula STOK_AKHIR belum dibuat.
 */
function buildProduksiSumMaps_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.OUTPUT_PRODUKSI);
  var maps = { byId: {}, byName: {} };
  if (!sheet || sheet.getLastRow() < 2) return maps;

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var tanggal = data[i][1];
    var id = asText_(data[i][3]);
    var nama = asText_(data[i][4]);
    var qty = parseNumber_(data[i][5]);
    var satuan = asText_(data[i][6]) || '-';
    if (!id && !nama) continue;

    var keyId = normalizeKey_(id);
    var keyName = normalizeKey_(nama);
    var item = null;

    if (keyId && maps.byId[keyId]) item = maps.byId[keyId];
    else if (keyName && maps.byName[keyName]) item = maps.byName[keyName];

    if (!item) {
      item = { id: id, nama: nama, qty: 0, satuan: satuan, lastUpdate: '' };
    }

    item.qty += qty;
    item.satuan = satuan || item.satuan;
    item.lastUpdate = maxDateText_(item.lastUpdate, formatSheetDate_(tanggal));

    if (keyId) maps.byId[keyId] = item;
    if (keyName) maps.byName[keyName] = item;
  }

  return maps;
}


/**
 * RIWAYAT v13/v14: gabungan transaksi untuk audit.
 * v14 ikut menampilkan Produk Outlet Terjual dan Opname Produk Outlet.
 */
function handleGetRiwayatTransaksi_(ss, requestData) {
  var profile = requestData._user || {};
  if (!profile.nama) {
    return dataResponse_({
      success: false,
      code: 'SESSION_EXPIRED',
      message: 'Sesi tidak valid atau user tidak terbaca. Silakan login ulang.'
    });
  }
  var userName = asText_(profile.nama);
  var level = asText_(profile.level);
  var jenisFilter = normalizeKey_(requestData.jenis || requestData.type || 'SEMUA');
  var keyword = normalizeKey_(requestData.keyword || requestData.search || '');
  var tanggalDari = asText_(requestData.tanggalDari || requestData.dateFrom || '');
  var tanggalSampai = asText_(requestData.tanggalSampai || requestData.dateTo || '');
  var isOwner = isOwnerRole_(level);
  var userOnly = !isOwner;
  var rows = [];

  addRiwayatFromSheet_(ss, rows, CONFIG.SHEETS.INPUT_STOK, 'Input', userName, userOnly, {
    tx: 0, tanggal: 1, petugas: 2, kategori: 3, id: 4, nama: 5, qty: 8, satuan: 9, tujuan: null, outlet: null, pic: null, catatan: null
  });
  addRiwayatFromSheet_(ss, rows, CONFIG.SHEETS.OUTPUT_STOK, 'Output', userName, userOnly, {
    tx: 0, tanggal: 1, petugas: 2, kategori: 3, id: 4, nama: 5, qty: 6, satuan: 7, tujuan: 8, outlet: 9, pic: 10, catatan: null
  });
  addRiwayatFromSheet_(ss, rows, CONFIG.SHEETS.OUTPUT_PRODUKSI, 'Produksi', userName, userOnly, {
    tx: 0, tanggal: 1, petugas: 2, kategori: null, id: 3, nama: 4, qty: 5, satuan: 6, tujuan: null, outlet: null, pic: null, catatan: null
  });
  addRiwayatFromSheet_(ss, rows, CONFIG.SHEETS.TRANSFER_PRODUKSI, 'Transfer', userName, userOnly, {
    tx: 0, tanggal: 1, petugas: 2, kategori: null, id: 5, nama: 6, qty: 7, satuan: 8, tujuan: 3, outlet: 3, pic: 4, catatan: 9
  });
  addRiwayatFromSheet_(ss, rows, CONFIG.SHEETS.STOK_OPNAME, 'Opname', userName, userOnly, {
    tx: 0, tanggal: 1, petugas: 2, kategori: 3, id: 4, nama: 5, qty: 7, satuan: 8, tujuan: null, outlet: null, pic: null, catatan: 9
  });
  addRiwayatFromSheet_(ss, rows, CONFIG.SHEETS.PRODUK_OUTLET_KELUAR, 'Produk Outlet Terjual', userName, userOnly, {
    tx: 0, tanggal: 1, outlet: 2, petugas: 3, kategori: null, id: 4, nama: 5, qty: 6, satuan: 7, tujuan: 2, pic: null, catatan: 8
  });
  addRiwayatFromSheet_(ss, rows, CONFIG.SHEETS.OPNAME_PRODUK_OUTLET, 'Opname Produk Outlet', userName, userOnly, {
    tx: 0, tanggal: 1, outlet: 2, petugas: 3, kategori: null, id: 4, nama: 5, qty: 8, satuan: 9, tujuan: 2, pic: null, catatan: 10
  });

  rows = rows.filter(function(row) {
    if (tanggalDari && row.tanggal < tanggalDari) return false;
    if (tanggalSampai && row.tanggal > tanggalSampai) return false;
    if (jenisFilter && jenisFilter !== 'semua' && jenisFilter !== 'all' && normalizeKey_(row.jenis) !== jenisFilter) return false;
    if (keyword) {
      var haystack = normalizeKey_([row.txId, row.jenis, row.tanggal, row.petugas, row.kategori, row.id, row.nama, row.qty, row.satuan, row.tujuan, row.outlet, row.pic, row.catatan].join(' '));
      if (haystack.indexOf(keyword) === -1) return false;
    }
    return true;
  });

  rows.sort(function(a, b) { return b.rank.localeCompare(a.rank); });
  rows.forEach(function(r) { delete r.rank; });

  return dataResponse_({ success: true, data: rows.slice(0, 1000), total: rows.length });
}

function addRiwayatFromSheet_(ss, target, sheetName, jenis, userName, userOnly, map) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var petugas = map.petugas === null || typeof map.petugas === 'undefined' ? '' : asText_(row[map.petugas]);
    if (userOnly && normalizeKey_(petugas) !== normalizeKey_(userName)) continue;
    var tanggal = formatSheetDate_(row[map.tanggal]);
    var txId = asText_(row[map.tx]);
    target.push({
      jenis: jenis,
      txId: txId,
      tanggal: tanggal,
      petugas: petugas,
      kategori: map.kategori === null || typeof map.kategori === 'undefined' ? jenis : asText_(row[map.kategori]),
      id: map.id === null || typeof map.id === 'undefined' ? '' : asText_(row[map.id]),
      nama: map.nama === null || typeof map.nama === 'undefined' ? '' : asText_(row[map.nama]),
      qty: map.qty === null || typeof map.qty === 'undefined' ? '' : asText_(row[map.qty]),
      satuan: map.satuan === null || typeof map.satuan === 'undefined' ? '' : asText_(row[map.satuan]),
      tujuan: map.tujuan === null || typeof map.tujuan === 'undefined' ? '' : asText_(row[map.tujuan]),
      outlet: map.outlet === null || typeof map.outlet === 'undefined' ? '' : asText_(row[map.outlet]),
      pic: map.pic === null || typeof map.pic === 'undefined' ? '' : asText_(row[map.pic]),
      catatan: map.catatan === null || typeof map.catatan === 'undefined' ? '' : asText_(row[map.catatan]),
      rank: buildActivityRank_(tanggal, txId, i)
    });
  }
}

/** PRODUK OUTLET v14: data stok outlet, terjual, dan opname. */
function handleGetProdukOutletData_(ss, requestData) {
  var tanggal = asText_(requestData.tanggal) || Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd');
  var profile = requestData._user;
  var petugas = asText_(profile && profile.nama);
  var requestedOutlet = asText_(requestData.outlet);
  if (!petugas) return dataResponse_({ success: false, message: 'Petugas tidak terbaca. Silakan logout lalu login ulang.' });

  if (!profile) return dataResponse_({ success: false, message: 'Data user tidak ditemukan di sheet USER.' });
  if (!isOwnerRole_(profile.level) && !permissionYes_(readPermissionsForLevel_(ss, profile.level), 'produkOutlet')) {
    return dataResponse_({ success: false, message: 'Akses ditolak. Role ini belum memiliki izin Produk Outlet.' });
  }

  var allowedOutlets = getAllowedOutletsForProfile_(ss, profile);
  if (!requestedOutlet) requestedOutlet = allowedOutlets[0] || '';
  if (!requestedOutlet) return dataResponse_({ success: false, message: 'Outlet user belum diisi di sheet USER.' });
  if (!isOutletAllowedForProfile_(profile, requestedOutlet)) return dataResponse_({ success: false, message: 'Akses outlet ditolak untuk user ini.' });

  var products = getMasterProdukOutlet_(ss, requestedOutlet);
  var moves = buildProdukOutletMovementMaps_(ss, requestedOutlet, tanggal);
  var rows = products.map(function(p) {
    var key = productKey_(p.id, p.nama);
    var stokKemarin = (moves.transferBefore[key] || 0) - (moves.keluarBefore[key] || 0) + (moves.opnameBefore[key] || 0);
    var stokMasuk = moves.transferToday[key] || 0;
    var stokTerjual = moves.keluarToday[key] || 0;
    var selisihOpname = moves.opnameToday[key] || 0;
    var stokSisa = stokKemarin + stokMasuk - stokTerjual + selisihOpname;
    return {
      id: p.id,
      nama: p.nama,
      satuan: p.satuan,
      urutan: p.urutan,
      keterangan: p.keterangan,
      stokKemarin: stokKemarin,
      stokMasuk: stokMasuk,
      stokTerjual: stokTerjual,
      selisihOpname: selisihOpname,
      stokSisa: stokSisa
    };
  });

  return dataResponse_({
    success: true,
    tanggal: tanggal,
    outlet: requestedOutlet,
    allowedOutlets: allowedOutlets,
    canSelectOutlet: isOwnerRole_(profile.level) || isAllOutlet_(profile.outlet),
    data: rows
  });
}

function handleSimpanProdukOutletKeluar_(ss, requestData) {
  var profile = requestData._user;
  var tanggal = asText_(requestData.tanggal);
  var outlet = asText_(requestData.outlet);
  var petugas = asText_(profile && profile.nama);
  var rows = requestData.rows || [];
  if (!tanggal) return dataResponse_({ success: false, message: 'Tanggal wajib diisi.' });
  if (!isValidDateText_(tanggal)) return dataResponse_({ success: false, message: 'Tanggal produk outlet tidak valid.' });
  if (!outlet) return dataResponse_({ success: false, message: 'Outlet wajib dipilih.' });
  if (!petugas) return dataResponse_({ success: false, message: 'Petugas tidak terbaca.' });

  if (!profile || (!isOwnerRole_(profile.level) && !permissionYes_(readPermissionsForLevel_(ss, profile.level), 'produkOutlet'))) return dataResponse_({ success: false, message: 'Akses Produk Outlet ditolak.' });
  if (!isOutletAllowedForProfile_(profile, outlet)) return dataResponse_({ success: false, message: 'Akses outlet ditolak.' });

  var current = handleGetProdukOutletDataRaw_(ss, tanggal, outlet);
  var currentByKey = {};
  current.forEach(function(r) { currentByKey[productKey_(r.id, r.nama)] = r; });

  var txId = buildTxId_('POT');
  var now = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');
  var output = [];
  rows.forEach(function(r) {
    var id = toSheetText_(r.id || r.idProduk);
    var nama = toSheetText_(r.nama || r.namaProduk);
    var qty = parseNumber_(r.qtyKeluar || r.qty || r.terjual);
    var satuan = toSheetText_(r.satuan) || '-';
    var ket = toSheetText_(r.keterangan || r.catatan || 'Terjual');
    if (!id && !nama) return;
    if (qty < 0) throw new Error('Qty terjual tidak boleh minus untuk ' + (nama || id) + '.');
    if (qty <= 0) return;
    var item = currentByKey[productKey_(id, nama)];
    var stokSisa = item ? parseNumber_(item.stokSisa) : 0;
    if (qty > stokSisa) throw new Error('Qty terjual ' + (nama || id) + ' melebihi stok sisa. Stok tersedia ' + stokSisa + ' ' + satuan + '.');
    output.push([txId, tanggal, outlet, petugas, id, nama, toSheetText_(qty), satuan, ket || 'Terjual', now]);
  });

  if (!output.length) return dataResponse_({ success: false, message: 'Tidak ada qty terjual yang diisi.' });
  return withScriptLock_(function() {
    var sheet = getOrCreateSheet_(ss, CONFIG.SHEETS.PRODUK_OUTLET_KELUAR, HEADER_PRODUK_OUTLET_KELUAR, false);
    sheet.getRange(sheet.getLastRow() + 1, 1, output.length, HEADER_PRODUK_OUTLET_KELUAR.length).setValues(output);
    appendAuditLog_(ss, requestData, 'SUCCESS', 'Produk outlet terjual disimpan. ID: ' + txId + '. Baris: ' + output.length + '.');
    return dataResponse_({ success: true, message: 'Produk terjual berhasil disimpan dengan ID: ' + txId });
  });
}

function handleSimpanProdukOutletOpname_(ss, requestData) {
  var profile = requestData._user;
  var tanggal = asText_(requestData.tanggal);
  var outlet = asText_(requestData.outlet);
  var petugas = asText_(profile && profile.nama);
  var rows = requestData.rows || [];
  if (!tanggal) return dataResponse_({ success: false, message: 'Tanggal wajib diisi.' });
  if (!isValidDateText_(tanggal)) return dataResponse_({ success: false, message: 'Tanggal opname produk outlet tidak valid.' });
  if (!outlet) return dataResponse_({ success: false, message: 'Outlet wajib dipilih.' });
  if (!petugas) return dataResponse_({ success: false, message: 'Petugas tidak terbaca.' });

  if (!profile || (!isOwnerRole_(profile.level) && !permissionYes_(readPermissionsForLevel_(ss, profile.level), 'produkOutlet'))) return dataResponse_({ success: false, message: 'Akses Produk Outlet ditolak.' });
  if (!isOutletAllowedForProfile_(profile, outlet)) return dataResponse_({ success: false, message: 'Akses outlet ditolak.' });

  var current = handleGetProdukOutletDataRaw_(ss, tanggal, outlet);
  var currentByKey = {};
  current.forEach(function(r) { currentByKey[productKey_(r.id, r.nama)] = r; });

  var opnId = buildTxId_('POP');
  var now = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss');
  var output = [];
  rows.forEach(function(r) {
    var id = toSheetText_(r.id || r.idProduk);
    var nama = toSheetText_(r.nama || r.namaProduk);
    var fisikRaw = r.stokFisik;
    if (fisikRaw === '' || fisikRaw === null || typeof fisikRaw === 'undefined') return;
    var stokFisik = parseNumber_(fisikRaw);
    if (stokFisik < 0) throw new Error('Stok fisik tidak boleh minus untuk ' + (nama || id) + '.');
    var satuan = toSheetText_(r.satuan) || '-';
    var alasan = toSheetText_(r.alasan || r.catatan);
    var item = currentByKey[productKey_(id, nama)];
    var stokSistem = item ? parseNumber_(item.stokSisa) : parseNumber_(r.stokSistem);
    var selisih = stokFisik - stokSistem;
    if (Math.abs(selisih) > 0.0001 && !alasan) throw new Error('Alasan opname wajib diisi untuk ' + (nama || id) + '.');
    output.push([opnId, tanggal, outlet, petugas, id, nama, toSheetText_(stokSistem), toSheetText_(stokFisik), toSheetText_(selisih), satuan, alasan, now]);
  });

  if (!output.length) return dataResponse_({ success: false, message: 'Tidak ada stok fisik yang diisi.' });
  return withScriptLock_(function() {
    var sheet = getOrCreateSheet_(ss, CONFIG.SHEETS.OPNAME_PRODUK_OUTLET, HEADER_OPNAME_PRODUK_OUTLET, false);
    sheet.getRange(sheet.getLastRow() + 1, 1, output.length, HEADER_OPNAME_PRODUK_OUTLET.length).setValues(output);
    appendAuditLog_(ss, requestData, 'SUCCESS', 'Opname produk outlet disimpan. ID: ' + opnId + '. Baris: ' + output.length + '.');
    return dataResponse_({ success: true, message: 'Opname produk outlet berhasil disimpan dengan ID: ' + opnId });
  });
}

function handleGetProdukOutletDataRaw_(ss, tanggal, outlet) {
  var products = getMasterProdukOutlet_(ss, outlet);
  var moves = buildProdukOutletMovementMaps_(ss, outlet, tanggal);
  return products.map(function(p) {
    var key = productKey_(p.id, p.nama);
    var stokKemarin = (moves.transferBefore[key] || 0) - (moves.keluarBefore[key] || 0) + (moves.opnameBefore[key] || 0);
    var stokMasuk = moves.transferToday[key] || 0;
    var stokTerjual = moves.keluarToday[key] || 0;
    var selisihOpname = moves.opnameToday[key] || 0;
    var stokSisa = stokKemarin + stokMasuk - stokTerjual + selisihOpname;
    return { id: p.id, nama: p.nama, satuan: p.satuan, stokKemarin: stokKemarin, stokMasuk: stokMasuk, stokTerjual: stokTerjual, selisihOpname: selisihOpname, stokSisa: stokSisa };
  });
}

function getMasterProdukOutlet_(ss, outlet) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER_PRODUK_OUTLET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var data = sheet.getDataRange().getValues();
  var headerMap = buildHeaderMap_(data[0]);
  var ixOutlet = findHeaderIndex_(headerMap, ['outlet'], 0);
  var ixId = findHeaderIndex_(headerMap, ['idproduk', 'idbarang', 'id'], 1);
  var ixNama = findHeaderIndex_(headerMap, ['namaproduk', 'produk', 'namabarang', 'nama'], 2);
  var ixSatuan = findHeaderIndex_(headerMap, ['satuan'], 3);
  var ixAktif = findHeaderIndex_(headerMap, ['aktif', 'active'], 4);
  var ixUrutan = findHeaderIndex_(headerMap, ['urutan', 'sort', 'no'], 5);
  var ixKet = findHeaderIndex_(headerMap, ['keterangan', 'catatan'], 6);
  var outKey = normalizeKey_(outlet);
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    if (normalizeKey_(data[i][ixOutlet]) !== outKey) continue;
    if (asYN_(data[i][ixAktif]) !== 'Y') continue;
    var id = asText_(data[i][ixId]);
    var nama = asText_(data[i][ixNama]);
    if (!id && !nama) continue;
    rows.push({ id: id, nama: nama, satuan: asText_(data[i][ixSatuan]) || '-', urutan: parseNumber_(data[i][ixUrutan]), keterangan: asText_(data[i][ixKet]) });
  }
  rows.sort(function(a, b) {
    if (a.urutan !== b.urutan) return a.urutan - b.urutan;
    return a.nama.localeCompare(b.nama);
  });
  return rows;
}

function buildProdukOutletMovementMaps_(ss, outlet, tanggal) {
  var maps = {
    transferBefore: {}, transferToday: {},
    keluarBefore: {}, keluarToday: {},
    opnameBefore: {}, opnameToday: {}
  };
  addProdukOutletTransferMaps_(ss, maps, outlet, tanggal);
  addProdukOutletKeluarMaps_(ss, maps, outlet, tanggal);
  addProdukOutletOpnameMaps_(ss, maps, outlet, tanggal);
  return maps;
}

function addMapQty_(map, key, qty) {
  if (!map[key]) map[key] = 0;
  map[key] += parseNumber_(qty);
}

function addProdukOutletTransferMaps_(ss, maps, outlet, tanggal) {
  var names = [CONFIG.SHEETS.TRANSFER_PRODUKSI, 'TRANSFER_PRODUKSI_STOK'];
  var seen = {};
  for (var n = 0; n < names.length; n++) {
    var sheetName = names[n];
    if (seen[sheetName]) continue;
    seen[sheetName] = true;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) continue;
    var data = sheet.getDataRange().getValues();
    var hm = buildHeaderMap_(data[0]);
    var ixTanggal = findHeaderIndex_(hm, ['tanggal'], 1);
    var ixOutlet = findHeaderIndex_(hm, ['outlet', 'outlettujuan'], 3);
    var ixId = findHeaderIndex_(hm, ['idproduk', 'idbarang', 'id'], 5);
    var ixNama = findHeaderIndex_(hm, ['namaproduk', 'produk', 'namabarang', 'nama'], 6);
    var ixQty = findHeaderIndex_(hm, ['qtytransfer', 'qtymasuk', 'qty'], 7);
    for (var i = 1; i < data.length; i++) {
      if (normalizeKey_(data[i][ixOutlet]) !== normalizeKey_(outlet)) continue;
      var d = formatSheetDate_(data[i][ixTanggal]);
      var key = productKey_(data[i][ixId], data[i][ixNama]);
      if (!key) continue;
      if (d < tanggal) addMapQty_(maps.transferBefore, key, data[i][ixQty]);
      else if (d === tanggal) addMapQty_(maps.transferToday, key, data[i][ixQty]);
    }
  }
}

function addProdukOutletKeluarMaps_(ss, maps, outlet, tanggal) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.PRODUK_OUTLET_KELUAR);
  if (!sheet || sheet.getLastRow() < 2) return;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeKey_(data[i][2]) !== normalizeKey_(outlet)) continue;
    var d = formatSheetDate_(data[i][1]);
    var key = productKey_(data[i][4], data[i][5]);
    if (!key) continue;
    if (d < tanggal) addMapQty_(maps.keluarBefore, key, data[i][6]);
    else if (d === tanggal) addMapQty_(maps.keluarToday, key, data[i][6]);
  }
}

function addProdukOutletOpnameMaps_(ss, maps, outlet, tanggal) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.OPNAME_PRODUK_OUTLET);
  if (!sheet || sheet.getLastRow() < 2) return;
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeKey_(data[i][2]) !== normalizeKey_(outlet)) continue;
    var d = formatSheetDate_(data[i][1]);
    var key = productKey_(data[i][4], data[i][5]);
    if (!key) continue;
    if (d < tanggal) addMapQty_(maps.opnameBefore, key, data[i][8]);
    else if (d === tanggal) addMapQty_(maps.opnameToday, key, data[i][8]);
  }
}

function productKey_(id, nama) {
  var idKey = normalizeKey_(id);
  if (idKey) return 'id|' + idKey;
  var nameKey = normalizeKey_(nama);
  return nameKey ? 'nama|' + nameKey : '';
}

function getUserProfileByName_(ss, userName) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.USER);
  if (!sheet || sheet.getLastRow() < 2) return null;
  var data = sheet.getDataRange().getValues();
  var hm = buildHeaderMap_(data[0]);
  var ixNama = findHeaderIndex_(hm, ['nama', 'namakaryawan', 'petugas'], 3);
  var ixLevel = findHeaderIndex_(hm, ['level', 'role', 'akses'], 4);
  var ixOutlet = findHeaderIndex_(hm, ['outlet', 'lokasi', 'cabang'], 5);
  for (var i = 1; i < data.length; i++) {
    if (normalizeKey_(data[i][ixNama]) === normalizeKey_(userName)) {
      var level = asText_(data[i][ixLevel]);
      var outlet = asText_(data[i][ixOutlet]);
      if (!outlet && isOwnerRole_(level)) outlet = 'ALL';
      return { nama: asText_(data[i][ixNama]), level: level, outlet: outlet };
    }
  }
  return null;
}

function getAllowedOutletsForProfile_(ss, profile) {
  if (!profile) return [];
  if (isOwnerRole_(profile.level) || isAllOutlet_(profile.outlet)) return getOutletNames_(ss);
  return profile.outlet ? [profile.outlet] : [];
}

function isAllOutlet_(outlet) {
  var o = normalizeKey_(outlet).replace(/\s+/g, '');
  return o === 'all' || o === 'alloutlet' || o === 'semua' || o === 'semuaoutlet';
}

function isOutletAllowedForProfile_(profile, outlet) {
  if (!profile) return false;
  if (isOwnerRole_(profile.level) || isAllOutlet_(profile.outlet)) return true;
  return normalizeKey_(profile.outlet) === normalizeKey_(outlet);
}

function withScriptLock_(callback) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function appendAuditLog_(ss, requestData, status, message) {
  try {
    requestData = requestData || {};
    var user = requestData._user || {};
    var sheet = getOrCreateSheet_(ss, CONFIG.SHEETS.AUDIT_LOG, HEADER_AUDIT_LOG, false);
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, HEADER_AUDIT_LOG.length).setValues([[
      Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy-MM-dd HH:mm:ss'),
      asText_(requestData.action),
      asText_(status),
      asText_(user.nama || user.username || requestData.userName || requestData.petugas || requestData.username),
      asText_(user.level),
      asText_(user.outlet || requestData.outlet),
      asText_(message).slice(0, 500),
      Utilities.getUuid()
    ]]);
  } catch (err) {
    // Audit log tidak boleh menggagalkan transaksi utama.
  }
}

/** Helper USER kolom D. */
function getUserNames_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEETS.USER);
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][3]) result.push(asText_(data[i][3]));
  }
  return uniqueSorted_(result);
}

/** Helper OUTLET: default membaca kolom A. Jika nama outlet ada di kolom B, ubah index 0 menjadi 1. */
function getOutletNames_(ss) {
  var result = [];
  var sheet = ss.getSheetByName(CONFIG.SHEETS.OUTLET);
  if (sheet && sheet.getLastRow() >= 2) {
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) result.push(asText_(data[i][0]));
    }
  }

  // Fallback: jika sheet OUTLET belum rapi, ambil nama outlet dari MASTER_PRODUK_OUTLET.
  var mpo = ss.getSheetByName(CONFIG.SHEETS.MASTER_PRODUK_OUTLET);
  if (mpo && mpo.getLastRow() >= 2) {
    var mdata = mpo.getDataRange().getValues();
    var hm = buildHeaderMap_(mdata[0]);
    var ixOutlet = findHeaderIndex_(hm, ['outlet'], 0);
    for (var r = 1; r < mdata.length; r++) {
      if (mdata[r][ixOutlet]) result.push(asText_(mdata[r][ixOutlet]));
    }
  }
  return uniqueSorted_(result);
}

function getOrCreateSheet_(ss, sheetName, header, forceHeader) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  // Untuk sheet produksi, header memang direvisi dari format lama menjadi format ringkas.
  if (forceHeader) {
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    var lastCol = sheet.getLastColumn();
    if (lastCol > header.length) {
      sheet.getRange(1, header.length + 1, 1, lastCol - header.length).clearContent();
    }
  }

  return sheet;
}

function buildTxId_(prefix) {
  var now = new Date();
  var dateStr = Utilities.formatDate(now, CONFIG.TZ, 'yyyyMMdd');
  var timeStr = Utilities.formatDate(now, CONFIG.TZ, 'HHmmss');
  var ms = ('000' + now.getMilliseconds()).slice(-3);
  var suffix = Utilities.getUuid().split('-')[0].toUpperCase();
  return prefix + '-' + dateStr + '-' + timeStr + ms + '-' + suffix;
}

function normalizeKey_(value) {
  return asText_(value).toLowerCase().trim();
}

function asText_(value) {
  if (value === null || typeof value === 'undefined') return '';
  return String(value).trim();
}

function asYN_(value) {
  var v = normalizeKey_(value);
  return (v === 'y' || v === 'ya' || v === 'yes' || v === 'true' || v === '1') ? 'Y' : 'N';
}

function passwordMatches_(passwordInput, storedPassword) {
  var password = asText_(passwordInput);
  var stored = asText_(storedPassword);
  if (!stored) return false;

  var parts = stored.split('$');
  if (parts.length === 2 && normalizeKey_(parts[0]) === 'sha256') {
    return sha256Hex_(password) === normalizeKey_(parts[1]);
  }
  if (parts.length === 3 && normalizeKey_(parts[0]) === 'sha256') {
    return sha256Hex_(parts[1] + password) === normalizeKey_(parts[2]);
  }

  // Kompatibilitas: password lama di sheet USER tetap bisa dipakai.
  return stored === password;
}

function sha256Hex_(text) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, asText_(text), Utilities.Charset.UTF_8);
  var out = '';
  for (var i = 0; i < bytes.length; i++) {
    var b = bytes[i];
    if (b < 0) b += 256;
    var h = b.toString(16);
    if (h.length < 2) h = '0' + h;
    out += h;
  }
  return out;
}

function buatHashPassword(password) {
  return 'sha256$' + sha256Hex_(password);
}

function buatHashPasswordDenganSalt(password, salt) {
  var saltText = asText_(salt) || Utilities.getUuid().split('-')[0];
  return 'sha256$' + saltText + '$' + sha256Hex_(saltText + asText_(password));
}

function toSheetText_(value) {
  if (value === null || typeof value === 'undefined') return '';
  // Mempertahankan gaya lama: decimal titik dikirim sebagai koma agar konsisten dengan format Indonesia.
  return String(value).replace('.', ',');
}

function parseNumber_(value) {
  if (typeof value === 'number') return isNaN(value) ? 0 : value;
  if (value === null || typeof value === 'undefined' || value === '') return 0;
  var cleaned = String(value).trim().replace(/\s/g, '');
  if (!cleaned) return 0;
  var commaIndex = cleaned.lastIndexOf(',');
  var dotIndex = cleaned.lastIndexOf('.');
  if (commaIndex !== -1 && dotIndex !== -1) {
    cleaned = commaIndex > dotIndex
      ? cleaned.replace(/\./g, '').replace(',', '.')
      : cleaned.replace(/,/g, '');
  } else if (commaIndex !== -1) {
    cleaned = cleaned.replace(',', '.');
  } else if ((cleaned.match(/\./g) || []).length > 1) {
    cleaned = cleaned.replace(/\./g, '');
  }
  var num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
}

function formatSheetDate_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, CONFIG.TZ, 'yyyy-MM-dd');
  }
  return asText_(value);
}

function maxDateText_(a, b) {
  if (!a) return b || '';
  if (!b) return a || '';
  return String(b) > String(a) ? b : a;
}

function uniqueSorted_(arr) {
  var seen = {};
  var out = [];
  arr.forEach(function(item) {
    var key = normalizeKey_(item);
    if (key && !seen[key]) {
      seen[key] = true;
      out.push(item);
    }
  });
  out.sort(function(a, b) { return a.localeCompare(b); });
  return out;
}

function dataResponse_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService
    .createTextOutput('API APJ Inventori v14.1 Transfer Produk Outlet Aktif.')
    .setMimeType(ContentService.MimeType.TEXT);
}
