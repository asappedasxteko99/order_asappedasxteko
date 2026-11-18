<script>
/**
 * ============================================================================
 * ASAP PEDAS X TEKO - BACKEND GOOGLE APPS SCRIPT (Code.gs)
 * Database: Google Spreadsheet ID 1lf71K9-c6ant6mI1LlSEJ3IAmLU8azzMY44Kc4ii9HM
 * Zona Waktu: Asia/Jakarta
 * ============================================================================
 */

const SPREADSHEET_ID = '1lf71K9-c6ant6mI1LlSEJ3IAmLU8azzMY44Kc4ii9HM';
const TIMEZONE = 'Asia/Jakarta';
const CACHE_TTL_SEC = 120;
const SESSION_HOURS = 12;

const ORDER_STATUSES = [
  'Menunggu Konfirmasi',
  'Dikonfirmasi',
  'Sedang Diproses',
  'Sedang Dimasak',
  'Siap Diambil / Dalam Pengiriman',
  'Selesai',
  'Dibatalkan'
];

const SCHEMA = {
  Settings: ['key', 'value', 'group', 'description', 'updated_at'],
  Users: ['id', 'username', 'full_name', 'role', 'password_hash', 'salt', 'is_active', 'created_at', 'updated_at', 'last_login'],
  Sessions: ['token', 'user_id', 'username', 'role', 'created_at', 'expires_at', 'revoked'],
  Products: ['id', 'category_id', 'name', 'description', 'image_url', 'price', 'promo_price', 'hpp', 'stock', 'is_available', 'is_popular', 'is_deleted', 'created_at', 'updated_at'],
  Categories: ['id', 'name', 'slug', 'sort_order', 'is_active', 'created_at', 'updated_at'],
  Variants: ['id', 'product_id', 'name', 'price_adjustment', 'is_active'],
  Banners: ['id', 'image_url', 'headline', 'subheadline', 'button_text', 'button_link', 'sort_order', 'is_active', 'created_at'],
  Orders: ['order_id', 'source', 'order_type', 'customer_id', 'customer_name', 'customer_phone', 'address', 'delivery_note', 'lat', 'lng', 'distance_km', 'distance_method', 'subtotal', 'discount_amount', 'promo_code', 'tax_amount', 'tax_percent', 'shipping_fee', 'total', 'payment_method', 'payment_status', 'paid_amount', 'change_amount', 'status', 'is_scheduled', 'scheduled_time', 'cashier_user_id', 'created_at', 'updated_at'],
  OrderItems: ['id', 'order_id', 'product_id', 'product_name_snapshot', 'variant_id', 'variant_name_snapshot', 'unit_price_snapshot', 'hpp_snapshot', 'qty', 'subtotal', 'note'],
  OrderStatusHistory: ['id', 'order_id', 'status', 'note', 'changed_by', 'changed_at'],
  Customers: ['id', 'name', 'phone', 'last_address', 'last_lat', 'last_lng', 'total_orders', 'total_spent', 'last_order_at', 'created_at', 'updated_at'],
  Transactions: ['trx_id', 'order_id', 'source', 'order_type', 'customer_id', 'customer_name', 'customer_phone', 'subtotal', 'discount_amount', 'tax_amount', 'shipping_fee', 'total', 'hpp_total', 'payment_method', 'paid_amount', 'change_amount', 'cashier_id', 'cashier_name', 'status', 'created_at'],
  TransactionItems: ['id', 'trx_id', 'order_id', 'product_id', 'product_name_snapshot', 'variant_name_snapshot', 'unit_price_snapshot', 'hpp_snapshot', 'qty', 'subtotal', 'note'],
  Expenses: ['id', 'date', 'category', 'amount', 'description', 'user_id', 'username', 'created_at'],
  Promotions: ['id', 'code', 'name', 'type', 'value', 'min_order', 'max_discount', 'start_date', 'end_date', 'is_active', 'created_at'],
  SocialMedia: ['id', 'platform', 'label', 'url', 'icon', 'type', 'sort_order', 'is_active'],
  DeliveryLocations: ['id', 'order_id', 'customer_phone', 'address', 'lat', 'lng', 'distance_km', 'route_provider', 'created_at'],
  Holidays: ['id', 'date', 'name', 'is_closed', 'note', 'created_at'],
  AuditLogs: ['id', 'user_id', 'username', 'role', 'action', 'target_type', 'target_id', 'before_data', 'after_data', 'created_at']
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ASAP PEDAS X TEKO — Pesan Online & POS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ============================================================================
 * UTILITIES & RESPONSE HELPERS
 * ============================================================================ */

function ok_(data, message) {
  return { ok: true, data: data || null, message: message || 'OK' };
}

function fail_(message, code) {
  return { ok: false, error: message || 'Terjadi kesalahan sistem.', code: code || 'ERR' };
}

function getSs_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function nowIso_() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

function todayDate_() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
}

function genId_(prefix) {
  return (prefix || 'ID') + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
}

function sanitizeStr_(val, maxLen) {
  if (val === null || val === undefined) return '';
  var s = String(val).trim();
  if (maxLen && s.length > maxLen) s = s.substring(0, maxLen);
  return s;
}

function toNum_(val, def) {
  var n = Number(val);
  return isNaN(n) ? (def || 0) : n;
}

function toBool_(val) {
  if (typeof val === 'boolean') return val;
  var s = String(val).toLowerCase().trim();
  return s === 'true' || s === '1' || s === 'yes';
}

function normalizePhone_(phone) {
  var p = String(phone || '').replace(/[^0-9]/g, '');
  if (p.indexOf('08') === 0) p = '62' + p.substring(1);
  if (p.indexOf('8') === 0) p = '62' + p;
  return p;
}

function getPepper_() {
  var props = PropertiesService.getScriptProperties();
  var pepper = props.getProperty('PASSWORD_PEPPER');
  if (!pepper) {
    pepper = Utilities.getUuid() + '-' + Utilities.getUuid();
    props.setProperty('PASSWORD_PEPPER', pepper);
  }
  return pepper;
}

function sha256Hex_(input) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  return bytes.map(function(b) {
    var v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function hashPassword_(password, salt) {
  var pepper = getPepper_();
  return sha256Hex_(salt + ':' + password + ':' + pepper);
}

function clearPublicCache_() {
  try {
    CacheService.getScriptCache().remove('PUBLIC_BOOTSTRAP_V1');
  } catch (e) {}
}

/* ============================================================================
 * DATABASE BATCH LAYER
 * ============================================================================ */

function getSheet_(name) {
  var ss = getSs_();
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" belum diinisialisasi. Jalankan setupDatabase() terlebih dahulu.');
  return sh;
}

function readAll_(sheetName) {
  var sh = getSheet_(sheetName);
  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0];
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var rowObj = { _rowIndex: i + 1 };
    var empty = true;
    for (var c = 0; c < headers.length; c++) {
      var h = headers[c];
      var v = values[i][c];
      if (v !== '' && v !== null && v !== undefined) empty = false;
      if (v instanceof Date) {
        v = Utilities.formatDate(v, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
      }
      rowObj[h] = v;
    }
    if (!empty) rows.push(rowObj);
  }
  return rows;
}

function appendRows_(sheetName, objArray) {
  if (!objArray || objArray.length === 0) return;
  var sh = getSheet_(sheetName);
  var headers = SCHEMA[sheetName];
  var rows = objArray.map(function(obj) {
    return headers.map(function(h) {
      return obj[h] !== undefined && obj[h] !== null ? obj[h] : '';
    });
  });
  var startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
}

function updateRowByIndex_(sheetName, rowIndex, obj) {
  var sh = getSheet_(sheetName);
  var headers = SCHEMA[sheetName];
  var row = headers.map(function(h) {
    return obj[h] !== undefined && obj[h] !== null ? obj[h] : '';
  });
  sh.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
}

function deleteRowByIndex_(sheetName, rowIndex) {
  var sh = getSheet_(sheetName);
  sh.deleteRow(rowIndex);
}

/* ============================================================================
 * SETUP DATABASE OTOMATIS (JALANKAN SEKALI)
 * ============================================================================ */

function setupDatabase() {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = getSs_();
    var sheetNames = Object.keys(SCHEMA);
    sheetNames.forEach(function(name) {
      var sh = ss.getSheetByName(name);
      if (!sh) {
        sh = ss.insertSheet(name);
      }
      var headers = SCHEMA[name];
      if (sh.getLastRow() === 0) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
        sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#111827').setFontColor('#ffffff');
        sh.setFrozenRows(1);
      } else {
        var existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
        if (!existing[0]) {
          sh.getRange(1, 1, 1, headers.length).setValues([headers]);
        }
      }
    });

    getPepper_();
    var now = nowIso_();

    // Seed Default Settings jika kosong
    var existingSettings = readAll_('Settings');
    if (existingSettings.length === 0) {
      var defaultSettings = [
        { key: 'restaurant_name', value: 'ASAP PEDAS X TEKO', group: 'identity', description: 'Nama restoran', updated_at: now },
        { key: 'logo_url', value: 'https://images.unsplash.com/photo-1555939594-58d7cb561ad1?w=300&q=80', group: 'identity', description: 'URL Logo restoran', updated_at: now },
        { key: 'address', value: 'Jl. Kuliner Nusantara No. 88, Jakarta Selatan', group: 'identity', description: 'Alamat restoran', updated_at: now },
        { key: 'whatsapp_number', value: '6281234567890', group: 'whatsapp', description: 'Nomor WhatsApp display', updated_at: now },
        { key: 'whatsapp_target', value: '6281234567890', group: 'whatsapp', description: 'Nomor WhatsApp tujuan order', updated_at: now },
        { key: 'whatsapp_template', value: 'Halo *ASAP PEDAS X TEKO*, saya ingin konfirmasi pesanan:\n\n*Order ID:* {{order_id}}\n*Nama:* {{name}}\n*No HP:* {{phone}}\n*Alamat:* {{address}}\n*Maps:* {{maps_url}}\n\n*Detail Pesanan:*\n{{items}}\n\nSubtotal: {{subtotal}}\nDiskon: {{discount}}\nPajak: {{tax}}\nOngkir ({{distance}} km): {{shipping}}\n*TOTAL: {{total}}*\n\nMohon konfirmasi pesanan dan kirimkan QRIS pembayaran. Terima kasih!', group: 'whatsapp', description: 'Template pesan WhatsApp', updated_at: now },
        { key: 'store_lat', value: '-6.229728', group: 'shipping', description: 'Latitude restoran', updated_at: now },
        { key: 'store_lng', value: '106.829518', group: 'shipping', description: 'Longitude restoran', updated_at: now },
        { key: 'tax_enabled', value: 'true', group: 'tax', description: 'Aktifkan pajak', updated_at: now },
        { key: 'tax_percent', value: '10', group: 'tax', description: 'Persentase pajak (%)', updated_at: now },
        { key: 'shipping_unit_km', value: '1', group: 'shipping', description: 'Satuan jarak dasar ongkir (KM)', updated_at: now },
        { key: 'shipping_rate_per_km', value: '3000', group: 'shipping', description: 'Tarif ongkir per satuan jarak KM', updated_at: now },
        { key: 'shipping_min_fee', value: '8000', group: 'shipping', description: 'Minimal ongkir', updated_at: now },
        { key: 'shipping_max_distance_km', value: '15', group: 'shipping', description: 'Maksimal jarak kirim (KM)', updated_at: now },
        { key: 'min_order_amount', value: '20000', group: 'shipping', description: 'Minimum order online', updated_at: now },
        { key: 'distance_method', value: 'GPS_HANDPHONE', group: 'shipping', description: 'Metode jarak GPS Handphone (Tanpa API)', updated_at: now },
        { key: 'road_factor', value: '1.2', group: 'shipping', description: 'Faktor koreksi jalan dari GPS (1.0 = garis lurus GPS)', updated_at: now },
        { key: 'manual_status', value: 'AUTO', group: 'operational', description: 'AUTO / OPEN / CLOSED', updated_at: now },
        { key: 'temp_closed', value: 'false', group: 'operational', description: 'Tutup sementara', updated_at: now },
        { key: 'temp_closed_reason', value: 'Sedang persiapan dapur', group: 'operational', description: 'Alasan tutup sementara', updated_at: now },
        { key: 'scheduled_order_enabled', value: 'true', group: 'operational', description: 'Izinkan pesanan terjadwal saat tutup', updated_at: now },
        { key: 'receipt_footer', value: 'Terima kasih telah menikmati sajian ASAP PEDAS X TEKO!\nPedasnya nendang, segarnya bikin nagih.', group: 'receipt', description: 'Footer struk', updated_at: now },
        { key: 'receipt_paper_size', value: '80mm', group: 'receipt', description: 'Ukuran kertas 58mm atau 80mm', updated_at: now },
        { key: 'qris_info', value: 'Pembayaran via QRIS / Transfer Bank akan dikonfirmasi admin via WhatsApp.', group: 'payment', description: 'Catatan QRIS', updated_at: now },
        { key: 'schedule_mon', value: '{"active":true,"open":"10:00","close":"22:00"}', group: 'operational', description: 'Jadwal Senin', updated_at: now },
        { key: 'schedule_tue', value: '{"active":true,"open":"10:00","close":"22:00"}', group: 'operational', description: 'Jadwal Selasa', updated_at: now },
        { key: 'schedule_wed', value: '{"active":true,"open":"10:00","close":"22:00"}', group: 'operational', description: 'Jadwal Rabu', updated_at: now },
        { key: 'schedule_thu', value: '{"active":true,"open":"10:00","close":"22:00"}', group: 'operational', description: 'Jadwal Kamis', updated_at: now },
        { key: 'schedule_fri', value: '{"active":true,"open":"10:00","close":"23:00"}', group: 'operational', description: 'Jadwal Jumat', updated_at: now },
        { key: 'schedule_sat', value: '{"active":true,"open":"10:00","close":"23:00"}', group: 'operational', description: 'Jadwal Sabtu', updated_at: now },
        { key: 'schedule_sun', value: '{"active":true,"open":"10:00","close":"22:00"}', group: 'operational', description: 'Jadwal Minggu', updated_at: now }
      ];
      appendRows_('Settings', defaultSettings);
    }

    // Seed Akun Admin Awal (admin / 123456)
    var existingUsers = readAll_('Users');
    if (existingUsers.length === 0) {
      var salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
      appendRows_('Users', [{
        id: 'USR-OWNER-01',
        username: 'admin',
        full_name: 'Owner ASAP PEDAS X TEKO',
        role: 'OWNER',
        password_hash: hashPassword_('123456', salt),
        salt: salt,
        is_active: true,
        created_at: now,
        updated_at: now,
        last_login: ''
      }]);
    }

    // Seed Kategori Awal
    var existingCats = readAll_('Categories');
    if (existingCats.length === 0) {
      appendRows_('Categories', [
        { id: 'CAT-01', name: 'Makanan', slug: 'makanan', sort_order: 1, is_active: true, created_at: now, updated_at: now },
        { id: 'CAT-02', name: 'Minuman', slug: 'minuman', sort_order: 2, is_active: true, created_at: now, updated_at: now },
        { id: 'CAT-03', name: 'Dessert', slug: 'dessert', sort_order: 3, is_active: true, created_at: now, updated_at: now }
      ]);
    }

    // Seed Produk Awal
    var existingProds = readAll_('Products');
    if (existingProds.length === 0) {
      appendRows_('Products', [
        { id: 'PRD-01', category_id: 'CAT-01', name: 'Ikan Asap Sambal Teko Level 1-5', description: 'Ikan asap gurih dipadu sambal teko pedas beraroma asap khas.', image_url: 'https://images.unsplash.com/photo-1544025162-d76694265947?w=600&q=80', price: 32000, promo_price: 29000, hpp: 16000, stock: 50, is_available: true, is_popular: true, is_deleted: false, created_at: now, updated_at: now },
        { id: 'PRD-02', category_id: 'CAT-01', name: 'Ayam Asap Pedas Nusantara', description: 'Ayam asap panggang bumbu rempah merah pedas mantap + lalapan.', image_url: 'https://images.unsplash.com/photo-1598515214211-89d3c73ae83b?w=600&q=80', price: 28000, promo_price: '', hpp: 14500, stock: 60, is_available: true, is_popular: true, is_deleted: false, created_at: now, updated_at: now },
        { id: 'PRD-03', category_id: 'CAT-01', name: 'Nasi Goreng Asap Iga', description: 'Nasi goreng smokey dengan potongan iga sapi empuk dan telur mata sapi.', image_url: 'https://images.unsplash.com/photo-1512058564366-18510be2db19?w=600&q=80', price: 38000, promo_price: '', hpp: 20000, stock: 40, is_available: true, is_popular: true, is_deleted: false, created_at: now, updated_at: now },
        { id: 'PRD-04', category_id: 'CAT-02', name: 'Es Teh Teko Jumbo', description: 'Teh seduh wangi melati disajikan dingin ukuran teko personal.', image_url: 'https://images.unsplash.com/photo-1556679343-c7306c1976bc?w=600&q=80', price: 10000, promo_price: '', hpp: 3000, stock: 100, is_available: true, is_popular: true, is_deleted: false, created_at: now, updated_at: now },
        { id: 'PRD-05', category_id: 'CAT-02', name: 'Es Jeruk Nipis Madu', description: 'Perasan jeruk nipis segar dengan madu murni penawar pedas.', image_url: 'https://images.unsplash.com/photo-1513558161293-cdaf765ed2fd?w=600&q=80', price: 14000, promo_price: '', hpp: 5000, stock: 80, is_available: true, is_popular: false, is_deleted: false, created_at: now, updated_at: now },
        { id: 'PRD-06', category_id: 'CAT-03', name: 'Puding Karamel Teko', description: 'Puding lembut saus karamel legit sebagai penutup pedas.', image_url: 'https://images.unsplash.com/photo-1488477181946-6428a0291777?w=600&q=80', price: 16000, promo_price: '', hpp: 6500, stock: 30, is_available: true, is_popular: false, is_deleted: false, created_at: now, updated_at: now }
      ]);
      appendRows_('Variants', [
        { id: 'VAR-01', product_id: 'PRD-01', name: 'Pedas Sedang', price_adjustment: 0, is_active: true },
        { id: 'VAR-02', product_id: 'PRD-01', name: 'Pedas Ekstra + Nasi', price_adjustment: 6000, is_active: true },
        { id: 'VAR-03', product_id: 'PRD-02', name: 'Paha Bawah', price_adjustment: 0, is_active: true },
        { id: 'VAR-04', product_id: 'PRD-02', name: 'Dada + Nasi Hangat', price_adjustment: 5000, is_active: true }
      ]);
    }

    // Seed Banners
    var existingBanners = readAll_('Banners');
    if (existingBanners.length === 0) {
      appendRows_('Banners', [
        { id: 'BNR-01', image_url: 'https://images.unsplash.com/photo-1555939594-58d7cb561ad1?w=1200&q=80', headline: 'SENSASI ASAP PEDAS X TEKO', subheadline: 'Ikan & Ayam Asap Sambal Teko autentik. Pesan langsung lebih hemat & cepat!', button_text: 'Pesan Sekarang', button_link: '#menu', sort_order: 1, is_active: true, created_at: now }
      ]);
    }

    // Seed SocialMedia
    var existingSocial = readAll_('SocialMedia');
    if (existingSocial.length === 0) {
      appendRows_('SocialMedia', [
        { id: 'SOC-01', platform: 'Instagram', label: '@asappedasxteko', url: 'https://instagram.com', icon: 'bi-instagram', type: 'SOCIAL', sort_order: 1, is_active: true },
        { id: 'SOC-02', platform: 'TikTok', label: '@asappedasxteko', url: 'https://tiktok.com', icon: 'bi-tiktok', type: 'SOCIAL', sort_order: 2, is_active: true },
        { id: 'SOC-03', platform: 'GrabFood', label: 'GrabFood', url: 'https://food.grab.com', icon: 'bi-bag-heart-fill', type: 'DELIVERY_APP', sort_order: 3, is_active: true },
        { id: 'SOC-04', platform: 'ShopeeFood', label: 'ShopeeFood', url: 'https://shopee.co.id', icon: 'bi-shop', type: 'DELIVERY_APP', sort_order: 4, is_active: true },
        { id: 'SOC-05', platform: 'GoFood', label: 'GoFood', url: 'https://gofood.co.id', icon: 'bi-scooter', type: 'DELIVERY_APP', sort_order: 5, is_active: true }
      ]);
    }

    // Seed Promo contoh
    var existingPromos = readAll_('Promotions');
    if (existingPromos.length === 0) {
      appendRows_('Promotions', [
        { id: 'PRM-01', code: 'TEKO10', name: 'Diskon 10% Pelanggan Online', type: 'PERCENT', value: 10, min_order: 50000, max_discount: 20000, start_date: '2025-01-01', end_date: '2028-12-31', is_active: true, created_at: now }
      ]);
    }

    clearPublicCache_();
    return ok_(null, 'Database berhasil diinisialisasi (19 sheet siap).');
  } catch (err) {
    return fail_('Gagal setupDatabase: ' + err.message);
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================================
 * SETTINGS & OPERATIONAL HOURS
 * ============================================================================ */

function getSettingsMap_() {
  var rows = readAll_('Settings');
  var map = {};
  rows.forEach(function(r) {
    map[r.key] = String(r.value);
  });
  return map;
}

function getOperationalStatus_(settingsMap, holidays) {
  var now = new Date();
  var todayStr = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd');
  var dayIndex = Number(Utilities.formatDate(now, TIMEZONE, 'u')); // 1=Mon..7=Sun
  var dayKeys = ['', 'schedule_mon', 'schedule_tue', 'schedule_wed', 'schedule_thu', 'schedule_fri', 'schedule_sat', 'schedule_sun'];
  var dayNames = ['', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'];
  var curTime = Utilities.formatDate(now, TIMEZONE, 'HH:mm');

  var manual = (settingsMap.manual_status || 'AUTO').toUpperCase();
  var tempClosed = toBool_(settingsMap.temp_closed);
  var scheduledEnabled = toBool_(settingsMap.scheduled_order_enabled);

  // Build weekly schedule summary
  var weekly = [];
  for (var d = 1; d <= 7; d++) {
    var cfg = { active: true, open: '10:00', close: '22:00' };
    try {
      if (settingsMap[dayKeys[d]]) cfg = JSON.parse(settingsMap[dayKeys[d]]);
    } catch (e) {}
    weekly.push({ dayIndex: d, dayName: dayNames[d], active: !!cfg.active, open: cfg.open || '10:00', close: cfg.close || '22:00' });
  }

  // Find next open info helper
  function findNextOpen() {
    for (var offset = 0; offset <= 7; offset++) {
      var checkIdx = ((dayIndex - 1 + offset) % 7) + 1;
      var checkCfg = weekly[checkIdx - 1];
      var checkDate = new Date(now.getTime() + offset * 86400000);
      var checkDateStr = Utilities.formatDate(checkDate, TIMEZONE, 'yyyy-MM-dd');
      var isHol = holidays.some(function(h) { return String(h.date).slice(0, 10) === checkDateStr && toBool_(h.is_closed); });
      if (!isHol && checkCfg.active) {
        if (offset === 0 && curTime < checkCfg.open) {
          return 'Buka hari ini pukul ' + checkCfg.open + ' WIB';
        }
        if (offset > 0) {
          return 'Buka kembali ' + (offset === 1 ? 'besok (' + checkCfg.dayName + ')' : checkCfg.dayName) + ' pukul ' + checkCfg.open + ' WIB';
        }
      }
    }
    return 'Cek jadwal operasional restoran';
  }

  if (tempClosed) {
    return {
      isOpen: false,
      statusLabel: 'CLOSED',
      reason: settingsMap.temp_closed_reason || 'Tutup Sementara',
      nextOpenInfo: findNextOpen(),
      scheduledOrderEnabled: scheduledEnabled,
      weeklySchedule: weekly
    };
  }

  if (manual === 'OPEN') {
    return {
      isOpen: true,
      statusLabel: 'OPEN',
      reason: 'Buka (Mode Manual)',
      nextOpenInfo: 'Sedang Buka Sekarang',
      scheduledOrderEnabled: scheduledEnabled,
      weeklySchedule: weekly
    };
  }
  if (manual === 'CLOSED') {
    return {
      isOpen: false,
      statusLabel: 'CLOSED',
      reason: 'Tutup (Mode Manual)',
      nextOpenInfo: findNextOpen(),
      scheduledOrderEnabled: scheduledEnabled,
      weeklySchedule: weekly
    };
  }

  var holidayToday = holidays.filter(function(h) {
    return String(h.date).slice(0, 10) === todayStr && toBool_(h.is_closed);
  })[0];
  if (holidayToday) {
    return {
      isOpen: false,
      statusLabel: 'CLOSED',
      reason: 'Libur: ' + holidayToday.name,
      nextOpenInfo: findNextOpen(),
      scheduledOrderEnabled: scheduledEnabled,
      weeklySchedule: weekly
    };
  }

  var todayCfg = weekly[dayIndex - 1];
  if (!todayCfg.active) {
    return {
      isOpen: false,
      statusLabel: 'CLOSED',
      reason: 'Hari ' + todayCfg.dayName + ' libur operasional',
      nextOpenInfo: findNextOpen(),
      scheduledOrderEnabled: scheduledEnabled,
      weeklySchedule: weekly
    };
  }

  var openNow = curTime >= todayCfg.open && curTime <= todayCfg.close;
  return {
    isOpen: openNow,
    statusLabel: openNow ? 'OPEN' : 'CLOSED',
    reason: openNow ? ('Buka s/d ' + todayCfg.close + ' WIB') : ('Di luar jam operasional (' + todayCfg.open + ' - ' + todayCfg.close + ' WIB)'),
    nextOpenInfo: openNow ? ('Buka sampai ' + todayCfg.close + ' WIB') : findNextOpen(),
    scheduledOrderEnabled: scheduledEnabled,
    weeklySchedule: weekly
  };
}

/* ============================================================================
 * AUTH, SESSION & RBAC + AUDIT LOG
 * ============================================================================ */

function writeAudit_(user, action, targetType, targetId, beforeObj, afterObj) {
  try {
    appendRows_('AuditLogs', [{
      id: genId_('AUD'),
      user_id: user ? user.id : 'SYSTEM',
      username: user ? user.username : 'SYSTEM',
      role: user ? user.role : 'SYSTEM',
      action: action,
      target_type: targetType || '',
      target_id: targetId || '',
      before_data: beforeObj ? JSON.stringify(beforeObj) : '',
      after_data: afterObj ? JSON.stringify(afterObj) : '',
      created_at: nowIso_()
    }]);
  } catch (e) {}
}

function requireAuth_(token, allowedRoles) {
  if (!token) throw new Error('Sesi tidak ditemukan. Silakan login kembali.');
  var sessions = readAll_('Sessions');
  var nowStr = nowIso_();
  var sess = null;
  for (var i = sessions.length - 1; i >= 0; i--) {
    if (sessions[i].token === token) {
      sess = sessions[i];
      break;
    }
  }
  if (!sess) throw new Error('Sesi tidak valid.');
  if (toBool_(sess.revoked)) throw new Error('Sesi telah dicabut. Silakan login kembali.');
  if (String(sess.expires_at) < nowStr) throw new Error('Sesi telah kedaluwarsa. Silakan login kembali.');

  var users = readAll_('Users');
  var user = users.filter(function(u) { return u.id === sess.user_id; })[0];
  if (!user || !toBool_(user.is_active)) throw new Error('Akun tidak aktif atau tidak ditemukan.');

  if (allowedRoles && allowedRoles.length > 0 && allowedRoles.indexOf(user.role) === -1) {
    throw new Error('Akses ditolak. Hak akses Anda (' + user.role + ') tidak diizinkan untuk aksi ini.');
  }
  return {
    id: user.id,
    username: user.username,
    fullName: user.full_name,
    role: user.role
  };
}

function loginUser(username, password) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var uname = sanitizeStr_(username, 50).toLowerCase();
    var pass = String(password || '');
    if (!uname || !pass) return fail_('Username dan password wajib diisi.');

    var users = readAll_('Users');
    var user = users.filter(function(u) { return String(u.username).toLowerCase() === uname; })[0];
    if (!user) return fail_('Username atau password salah.');
    if (!toBool_(user.is_active)) return fail_('Akun dinonaktifkan oleh administrator.');

    var computed = hashPassword_(pass, String(user.salt));
    if (computed !== String(user.password_hash)) {
      writeAudit_({ id: user.id, username: user.username, role: user.role }, 'LOGIN_FAILED', 'AUTH', user.id, null, { reason: 'Password salah' });
      return fail_('Username atau password salah.');
    }

    var now = new Date();
    var createdAt = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    var expiresAt = Utilities.formatDate(new Date(now.getTime() + SESSION_HOURS * 3600000), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
    var token = sha256Hex_(Utilities.getUuid() + ':' + user.id + ':' + now.getTime());

    appendRows_('Sessions', [{
      token: token,
      user_id: user.id,
      username: user.username,
      role: user.role,
      created_at: createdAt,
      expires_at: expiresAt,
      revoked: false
    }]);

    user.last_login = createdAt;
    updateRowByIndex_('Users', user._rowIndex, user);

    var safeUser = { id: user.id, username: user.username, fullName: user.full_name, role: user.role };
    writeAudit_(safeUser, 'LOGIN_SUCCESS', 'AUTH', user.id, null, { expires_at: expiresAt });

    return ok_({ token: token, user: safeUser, expiresAt: expiresAt }, 'Login berhasil.');
  } catch (err) {
    return fail_(err.message);
  } finally {
    lock.releaseLock();
  }
}

function logoutUser(token) {
  try {
    if (!token) return ok_(null);
    var sessions = readAll_('Sessions');
    for (var i = sessions.length - 1; i >= 0; i--) {
      if (sessions[i].token === token) {
        var s = sessions[i];
        s.revoked = true;
        updateRowByIndex_('Sessions', s._rowIndex, s);
        writeAudit_({ id: s.user_id, username: s.username, role: s.role }, 'LOGOUT', 'AUTH', s.user_id, null, null);
        break;
      }
    }
    return ok_(null, 'Berhasil logout.');
  } catch (err) {
    return fail_(err.message);
  }
}

function verifySession(token) {
  try {
    var user = requireAuth_(token, null);
    return ok_({ user: user });
  } catch (err) {
    return fail_(err.message, 'UNAUTHORIZED');
  }
}

/* ============================================================================
 * PUBLIC STOREFRONT DATA
 * ============================================================================ */

function getPublicBootstrap() {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get('PUBLIC_BOOTSTRAP_V1');
    var data;
    if (cached) {
      data = JSON.parse(cached);
      // Re-compute operational time status on every request so OPEN/CLOSED is exact
      var settingsMap = getSettingsMap_();
      var holidays = readAll_('Holidays');
      data.operational = getOperationalStatus_(settingsMap, holidays);
      return ok_(data);
    }

    var settingsMap = getSettingsMap_();
    var categories = readAll_('Categories')
      .filter(function(c) { return toBool_(c.is_active); })
      .sort(function(a, b) { return toNum_(a.sort_order) - toNum_(b.sort_order); });

    var variants = readAll_('Variants').filter(function(v) { return toBool_(v.is_active); });
    var varByProd = {};
    variants.forEach(function(v) {
      if (!varByProd[v.product_id]) varByProd[v.product_id] = [];
      varByProd[v.product_id].push({
        id: v.id,
        name: v.name,
        price_adjustment: toNum_(v.price_adjustment, 0)
      });
    });

    var products = readAll_('Products')
      .filter(function(p) { return !toBool_(p.is_deleted); })
      .map(function(p) {
        var basePrice = toNum_(p.price, 0);
        var promoPrice = p.promo_price !== '' && toNum_(p.promo_price, 0) > 0 ? toNum_(p.promo_price, 0) : null;
        return {
          id: p.id,
          category_id: p.category_id,
          name: p.name,
          description: p.description,
          image_url: p.image_url,
          price: basePrice,
          promo_price: promoPrice,
          effective_price: promoPrice !== null && promoPrice < basePrice ? promoPrice : basePrice,
          stock: toNum_(p.stock, 0),
          is_available: toBool_(p.is_available) && toNum_(p.stock, 0) > 0,
          is_popular: toBool_(p.is_popular),
          variants: varByProd[p.id] || []
        };
      });

    var banners = readAll_('Banners')
      .filter(function(b) { return toBool_(b.is_active); })
      .sort(function(a, b) { return toNum_(a.sort_order) - toNum_(b.sort_order); });

    var socialMedia = readAll_('SocialMedia')
      .filter(function(s) { return toBool_(s.is_active); })
      .sort(function(a, b) { return toNum_(a.sort_order) - toNum_(b.sort_order); });

    var today = todayDate_();
    var promotions = readAll_('Promotions')
      .filter(function(pr) {
        return toBool_(pr.is_active) && String(pr.start_date).slice(0, 10) <= today && String(pr.end_date).slice(0, 10) >= today;
      })
      .map(function(pr) {
        return {
          code: pr.code,
          name: pr.name,
          type: pr.type,
          value: toNum_(pr.value, 0),
          min_order: toNum_(pr.min_order, 0),
          max_discount: toNum_(pr.max_discount, 0)
        };
      });

    var publicSettings = {
      restaurant_name: settingsMap.restaurant_name || 'ASAP PEDAS X TEKO',
      logo_url: settingsMap.logo_url || '',
      address: settingsMap.address || '',
      whatsapp_number: settingsMap.whatsapp_number || '',
      store_lat: toNum_(settingsMap.store_lat, -6.229728),
      store_lng: toNum_(settingsMap.store_lng, 106.829518),
      tax_enabled: toBool_(settingsMap.tax_enabled),
      tax_percent: toNum_(settingsMap.tax_percent, 0),
      shipping_unit_km: Math.max(0.1, toNum_(settingsMap.shipping_unit_km, 1)),
      shipping_rate_per_km: toNum_(settingsMap.shipping_rate_per_km, 3000),
      shipping_min_fee: toNum_(settingsMap.shipping_min_fee, 8000),
      shipping_max_distance_km: toNum_(settingsMap.shipping_max_distance_km, 15),
      min_order_amount: toNum_(settingsMap.min_order_amount, 0),
      distance_method: settingsMap.distance_method || 'GPS_HANDPHONE',
      qris_info: settingsMap.qris_info || ''
    };

    var holidays = readAll_('Holidays');
    data = {
      settings: publicSettings,
      operational: getOperationalStatus_(settingsMap, holidays),
      categories: categories,
      products: products,
      banners: banners,
      socialMedia: socialMedia,
      promotions: promotions
    };

    try {
      cache.put('PUBLIC_BOOTSTRAP_V1', JSON.stringify(data), CACHE_TTL_SEC);
    } catch (e) {}

    return ok_(data);
  } catch (err) {
    return fail_(err.message);
  }
}

/* ============================================================================
 * ROUTING DISTANCE & BACKEND PRICING ENGINE
 * ============================================================================ */

function computeRouteDistanceKm_(storeLat, storeLng, custLat, custLng, roadFactorVal) {
  if (custLat === null || custLng === null || custLat === '' || custLng === '' || isNaN(custLat) || isNaN(custLng)) {
    throw new Error('Koordinat GPS Handphone belum terisi. Tekan tombol Gunakan GPS Saya terlebih dahulu.');
  }
  if (Math.abs(custLat) > 90 || Math.abs(custLng) > 180) {
    throw new Error('Koordinat GPS berada di luar rentang valid.');
  }

  // Hitung jarak langsung dari koordinat GPS Handphone tanpa API eksternal
  var toRad = function(deg) { return deg * Math.PI / 180; };
  var R = 6371; // Radius bumi (km)
  var dLat = toRad(custLat - storeLat);
  var dLng = toRad(custLng - storeLng);
  var a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(toRad(storeLat)) * Math.cos(toRad(custLat)) *
    Math.sin(dLng / 2) * Math.sin(dLng / 2);
  var c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  var straightKm = R * c;

  var factor = toNum_(roadFactorVal, 1.2);
  if (factor < 1) factor = 1;
  var estimatedKm = Math.round((straightKm * factor) * 100) / 100;

  return {
    km: estimatedKm,
    provider: 'GPS Handphone (Tanpa API)'
  };
}

function buildValidatedOrderCalc_(payload, isOnline) {
  var settingsMap = getSettingsMap_();
  var products = readAll_('Products');
  var variants = readAll_('Variants');

  var prodMap = {};
  products.forEach(function(p) {
    if (!toBool_(p.is_deleted)) prodMap[p.id] = p;
  });
  var varMap = {};
  variants.forEach(function(v) {
    if (toBool_(v.is_active)) varMap[v.id] = v;
  });

  if (!payload || !payload.items || !payload.items.length) {
    throw new Error('Keranjang masih kosong.');
  }

  var validatedItems = [];
  var subtotal = 0;
  var hppTotal = 0;

  payload.items.forEach(function(item) {
    var prod = prodMap[item.productId];
    if (!prod) throw new Error('Produk tidak ditemukan atau telah dihapus.');
    if (!toBool_(prod.is_available) || toNum_(prod.stock, 0) <= 0) {
      throw new Error('Produk "' + prod.name + '" sedang habis.');
    }
    var qty = parseInt(item.qty, 10);
    if (isNaN(qty) || qty < 1 || qty > 500) throw new Error('Jumlah item tidak valid.');
    if (qty > toNum_(prod.stock, 0)) {
      throw new Error('Stok "' + prod.name + '" tersisa ' + prod.stock + '.');
    }

    var basePrice = toNum_(prod.price, 0);
    var promoPrice = prod.promo_price !== '' && toNum_(prod.promo_price, 0) > 0 ? toNum_(prod.promo_price, 0) : null;
    var unitPrice = promoPrice !== null && promoPrice < basePrice ? promoPrice : basePrice;

    var varObj = null;
    if (item.variantId) {
      varObj = varMap[item.variantId];
      if (!varObj || varObj.product_id !== prod.id) {
        throw new Error('Variasi untuk "' + prod.name + '" tidak valid.');
      }
      unitPrice += toNum_(varObj.price_adjustment, 0);
    }

    var lineSubtotal = unitPrice * qty;
    var itemHpp = toNum_(prod.hpp, 0);
    subtotal += lineSubtotal;
    hppTotal += itemHpp * qty;

    validatedItems.push({
      productId: prod.id,
      productName: prod.name,
      variantId: varObj ? varObj.id : '',
      variantName: varObj ? varObj.name : '',
      unitPrice: unitPrice,
      hpp: itemHpp,
      qty: qty,
      subtotal: lineSubtotal,
      note: sanitizeStr_(item.note, 200)
    });
  });

  // Promo / Discount Validation
  var discountAmount = 0;
  var promoCodeUsed = '';
  var promoCodeInput = sanitizeStr_(payload.promoCode, 40).toUpperCase();
  if (promoCodeInput) {
    var today = todayDate_();
    var promos = readAll_('Promotions');
    var promo = promos.filter(function(pr) {
      return String(pr.code).toUpperCase() === promoCodeInput && toBool_(pr.is_active);
    })[0];
    if (!promo) throw new Error('Kode promo "' + promoCodeInput + '" tidak ditemukan atau tidak aktif.');
    if (String(promo.start_date).slice(0, 10) > today || String(promo.end_date).slice(0, 10) < today) {
      throw new Error('Kode promo "' + promoCodeInput + '" sudah kedaluwarsa atau belum berlaku.');
    }
    if (subtotal < toNum_(promo.min_order, 0)) {
      throw new Error('Minimal belanja untuk promo ini adalah Rp ' + toNum_(promo.min_order, 0).toLocaleString('id-ID') + '.');
    }
    if (promo.type === 'PERCENT') {
      discountAmount = Math.round(subtotal * (toNum_(promo.value, 0) / 100));
      var maxDisc = toNum_(promo.max_discount, 0);
      if (maxDisc > 0 && discountAmount > maxDisc) discountAmount = maxDisc;
    } else {
      discountAmount = toNum_(promo.value, 0);
    }
    if (discountAmount > subtotal) discountAmount = subtotal;
    promoCodeUsed = promo.code;
  } else if (!isOnline && toNum_(payload.manualDiscount, 0) > 0) {
    // POS manual discount
    discountAmount = Math.min(subtotal, Math.max(0, Math.round(toNum_(payload.manualDiscount, 0))));
  }

  var afterDiscount = Math.max(0, subtotal - discountAmount);

  // Tax calculation
  var taxEnabled = toBool_(settingsMap.tax_enabled);
  var taxPercent = taxEnabled ? toNum_(settingsMap.tax_percent, 0) : 0;
  var taxAmount = taxEnabled ? Math.round(afterDiscount * (taxPercent / 100)) : 0;

  // Shipping & Distance Calculation
  var distanceKm = 0;
  var distanceProvider = '';
  var shippingFee = 0;
  var orderType = sanitizeStr_(payload.orderType || (isOnline ? 'DELIVERY' : 'DINE_IN'), 30).toUpperCase();
  var minOrderAmount = toNum_(settingsMap.min_order_amount, 0);
  var maxDistKm = toNum_(settingsMap.shipping_max_distance_km, 15);

  if (isOnline) {
    if (subtotal < minOrderAmount) {
      throw new Error('Minimum order adalah Rp ' + minOrderAmount.toLocaleString('id-ID') + '.');
    }
    if (orderType === 'DELIVERY') {
      var storeLat = toNum_(settingsMap.store_lat, -6.229728);
      var storeLng = toNum_(settingsMap.store_lng, 106.829518);
      var custLat = Number(payload.lat);
      var custLng = Number(payload.lng);
      var routeInfo = computeRouteDistanceKm_(storeLat, storeLng, custLat, custLng, settingsMap.road_factor);
      distanceKm = routeInfo.km;
      distanceProvider = routeInfo.provider;

      if (maxDistKm > 0 && distanceKm > maxDistKm) {
        throw new Error('Jarak pengiriman (' + distanceKm + ' km) melebihi batas maksimal restoran (' + maxDistKm + ' km).');
      }
      var unitKm = Math.max(0.1, toNum_(settingsMap.shipping_unit_km, 1));
      var ratePerUnit = Math.max(0, toNum_(settingsMap.shipping_rate_per_km, 3000));
      var minFee = Math.max(0, toNum_(settingsMap.shipping_min_fee, 8000));
      shippingFee = Math.max(minFee, Math.round((distanceKm / unitKm) * ratePerUnit));
    }
  } else {
    // POS manual shipping if Delivery at POS
    if (orderType === 'DELIVERY' && toNum_(payload.manualShippingFee, 0) > 0) {
      shippingFee = Math.round(toNum_(payload.manualShippingFee, 0));
    }
  }

  var total = afterDiscount + taxAmount + shippingFee;

  return {
    settingsMap: settingsMap,
    items: validatedItems,
    subtotal: subtotal,
    discountAmount: discountAmount,
    promoCode: promoCodeUsed,
    taxEnabled: taxEnabled,
    taxPercent: taxPercent,
    taxAmount: taxAmount,
    distanceKm: distanceKm,
    distanceProvider: distanceProvider,
    shippingFee: shippingFee,
    total: total,
    hppTotal: hppTotal
  };
}

function calculateOrderQuote(payload) {
  try {
    var calc = buildValidatedOrderCalc_(payload, true);
    return ok_({
      items: calc.items,
      subtotal: calc.subtotal,
      discountAmount: calc.discountAmount,
      promoCode: calc.promoCode,
      taxEnabled: calc.taxEnabled,
      taxPercent: calc.taxPercent,
      taxAmount: calc.taxAmount,
      distanceKm: calc.distanceKm,
      distanceProvider: calc.distanceProvider,
      shippingFee: calc.shippingFee,
      total: calc.total
    });
  } catch (err) {
    return fail_(err.message);
  }
}

/* ============================================================================
 * ONLINE CHECKOUT & WHATSAPP ORDER + TRACKING
 * ============================================================================ */

function generateUniqueOrderId_() {
  var datePart = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd');
  var existing = readAll_('Orders').map(function(o) { return o.order_id; });
  for (var attempt = 0; attempt < 10; attempt++) {
    var rand = Math.floor(1000 + Math.random() * 9000);
    var candidate = 'ORD-' + datePart + '-' + rand;
    if (existing.indexOf(candidate) === -1) return candidate;
  }
  return 'ORD-' + datePart + '-' + Date.now().toString().slice(-5);
}

function upsertCustomer_(name, phone, address, lat, lng, orderTotal, now) {
  var normPhone = normalizePhone_(phone);
  var customers = readAll_('Customers');
  var found = customers.filter(function(c) { return normalizePhone_(c.phone) === normPhone; })[0];
  if (found) {
    found.name = name || found.name;
    if (address) found.last_address = address;
    if (lat !== '' && lat !== undefined) found.last_lat = lat;
    if (lng !== '' && lng !== undefined) found.last_lng = lng;
    found.total_orders = toNum_(found.total_orders, 0) + 1;
    found.total_spent = toNum_(found.total_spent, 0) + toNum_(orderTotal, 0);
    found.last_order_at = now;
    found.updated_at = now;
    updateRowByIndex_('Customers', found._rowIndex, found);
    return found.id;
  } else {
    var newId = genId_('CUS');
    appendRows_('Customers', [{
      id: newId,
      name: name,
      phone: normPhone,
      last_address: address || '',
      last_lat: lat || '',
      last_lng: lng || '',
      total_orders: 1,
      total_spent: toNum_(orderTotal, 0),
      last_order_at: now,
      created_at: now,
      updated_at: now
    }]);
    return newId;
  }
}

function placeOnlineOrder(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var settingsMap = getSettingsMap_();
    var holidays = readAll_('Holidays');
    var op = getOperationalStatus_(settingsMap, holidays);

    var isScheduled = toBool_(payload.isScheduled);
    if (!op.isOpen) {
      if (!op.scheduledOrderEnabled) {
        return fail_('Restoran sedang tutup (' + op.reason + '). ' + op.nextOpenInfo);
      }
      if (!isScheduled || !sanitizeStr_(payload.scheduledTime, 40)) {
        return fail_('Restoran sedang tutup. Aktifkan pesanan terjadwal dan pilih waktu pengiriman/pengambilan.');
      }
    }

    var custName = sanitizeStr_(payload.customerName, 100);
    var custPhone = normalizePhone_(payload.customerPhone);
    var address = sanitizeStr_(payload.address, 400);
    var deliveryNote = sanitizeStr_(payload.deliveryNote, 250);
    var orderType = sanitizeStr_(payload.orderType || 'DELIVERY', 30).toUpperCase();
    var paymentMethod = sanitizeStr_(payload.paymentMethod || 'QRIS / Transfer', 50);

    if (!custName) return fail_('Nama pelanggan wajib diisi.');
    if (!custPhone || custPhone.length < 9) return fail_('Nomor telepon/WhatsApp wajib diisi dengan benar.');
    if (orderType === 'DELIVERY' && !address) return fail_('Alamat lengkap wajib diisi.');
    if (paymentMethod.toUpperCase().indexOf('COD') !== -1) {
      return fail_('Maaf, kami tidak menerima pembayaran COD.');
    }

    var calc = buildValidatedOrderCalc_(payload, true);
    var orderId = generateUniqueOrderId_();
    var now = nowIso_();

    var customerId = upsertCustomer_(custName, custPhone, address, payload.lat, payload.lng, calc.total, now);

    // Kurangi stok produk secara batch
    var allProds = readAll_('Products');
    calc.items.forEach(function(item) {
      var p = allProds.filter(function(pr) { return pr.id === item.productId; })[0];
      if (p) {
        p.stock = Math.max(0, toNum_(p.stock, 0) - item.qty);
        if (p.stock === 0) p.is_available = false;
        p.updated_at = now;
        updateRowByIndex_('Products', p._rowIndex, p);
      }
    });

    appendRows_('Orders', [{
      order_id: orderId,
      source: 'ONLINE',
      order_type: orderType,
      customer_id: customerId,
      customer_name: custName,
      customer_phone: custPhone,
      address: address,
      delivery_note: deliveryNote,
      lat: payload.lat || '',
      lng: payload.lng || '',
      distance_km: calc.distanceKm,
      distance_method: calc.distanceProvider,
      subtotal: calc.subtotal,
      discount_amount: calc.discountAmount,
      promo_code: calc.promoCode,
      tax_amount: calc.taxAmount,
      tax_percent: calc.taxPercent,
      shipping_fee: calc.shippingFee,
      total: calc.total,
      payment_method: paymentMethod,
      payment_status: 'UNPAID',
      paid_amount: 0,
      change_amount: 0,
      status: 'Menunggu Konfirmasi',
      is_scheduled: isScheduled,
      scheduled_time: sanitizeStr_(payload.scheduledTime, 60),
      cashier_user_id: '',
      created_at: now,
      updated_at: now
    }]);

    var orderItemRows = calc.items.map(function(it) {
      return {
        id: genId_('OIT'),
        order_id: orderId,
        product_id: it.productId,
        product_name_snapshot: it.productName,
        variant_id: it.variantId,
        variant_name_snapshot: it.variantName,
        unit_price_snapshot: it.unitPrice,
        hpp_snapshot: it.hpp,
        qty: it.qty,
        subtotal: it.subtotal,
        note: it.note
      };
    });
    appendRows_('OrderItems', orderItemRows);

    appendRows_('OrderStatusHistory', [{
      id: genId_('OSH'),
      order_id: orderId,
      status: 'Menunggu Konfirmasi',
      note: 'Pesanan online dibuat oleh pelanggan. Menunggu konfirmasi & pembayaran QRIS/Transfer.',
      changed_by: 'PELANGGAN',
      changed_at: now
    }]);

    if (orderType === 'DELIVERY' && payload.lat && payload.lng) {
      appendRows_('DeliveryLocations', [{
        id: genId_('LOC'),
        order_id: orderId,
        customer_phone: custPhone,
        address: address,
        lat: payload.lat,
        lng: payload.lng,
        distance_km: calc.distanceKm,
        route_provider: calc.distanceProvider,
        created_at: now
      }]);
    }

    clearPublicCache_();

    // Build WhatsApp message
    var itemsText = calc.items.map(function(it, idx) {
      var line = (idx + 1) + '. ' + it.productName +
        (it.variantName ? ' (' + it.variantName + ')' : '') +
        ' x' + it.qty + ' @ Rp ' + it.unitPrice.toLocaleString('id-ID') +
        ' = Rp ' + it.subtotal.toLocaleString('id-ID');
      if (it.note) line += '\n   Catatan: ' + it.note;
      return line;
    }).join('\n');

    var mapsUrl = (payload.lat && payload.lng)
      ? ('https://maps.google.com/?q=' + payload.lat + ',' + payload.lng)
      : '-';

    var tpl = settingsMap.whatsapp_template ||
      'Halo *ASAP PEDAS X TEKO*, konfirmasi pesanan:\nOrder ID: {{order_id}}\nNama: {{name}}\nHP: {{phone}}\nAlamat: {{address}}\nMaps: {{maps_url}}\n\n{{items}}\n\nSubtotal: {{subtotal}}\nDiskon: {{discount}}\nPajak: {{tax}}\nOngkir: {{shipping}}\nTotal: {{total}}\n\nMohon konfirmasi pesanan & QRIS pembayaran.';

    var waText = tpl
      .replace(/\{\{order_id\}\}/g, orderId)
      .replace(/\{\{name\}\}/g, custName)
      .replace(/\{\{phone\}\}/g, custPhone)
      .replace(/\{\{address\}\}/g, address + (deliveryNote ? ' (' + deliveryNote + ')' : ''))
      .replace(/\{\{maps_url\}\}/g, mapsUrl)
      .replace(/\{\{items\}\}/g, itemsText)
      .replace(/\{\{subtotal\}\}/g, 'Rp ' + calc.subtotal.toLocaleString('id-ID'))
      .replace(/\{\{discount\}\}/g, 'Rp ' + calc.discountAmount.toLocaleString('id-ID'))
      .replace(/\{\{tax\}\}/g, 'Rp ' + calc.taxAmount.toLocaleString('id-ID'))
      .replace(/\{\{distance\}\}/g, String(calc.distanceKm))
      .replace(/\{\{shipping\}\}/g, 'Rp ' + calc.shippingFee.toLocaleString('id-ID'))
      .replace(/\{\{total\}\}/g, 'Rp ' + calc.total.toLocaleString('id-ID'));

    if (isScheduled && payload.scheduledTime) {
      waText += '\n*Jadwal Pengiriman/Ambil:* ' + sanitizeStr_(payload.scheduledTime, 60);
    }

    var targetWa = normalizePhone_(settingsMap.whatsapp_target || settingsMap.whatsapp_number || '6281234567890');
    var waUrl = 'https://wa.me/' + targetWa + '?text=' + encodeURIComponent(waText);

    return ok_({
      orderId: orderId,
      status: 'Menunggu Konfirmasi',
      total: calc.total,
      quote: {
        subtotal: calc.subtotal,
        discountAmount: calc.discountAmount,
        taxAmount: calc.taxAmount,
        shippingFee: calc.shippingFee,
        distanceKm: calc.distanceKm,
        distanceProvider: calc.distanceProvider
      },
      whatsappUrl: waUrl,
      whatsappText: waText
    }, 'Pesanan berhasil disimpan dengan ID ' + orderId);
  } catch (err) {
    return fail_(err.message);
  } finally {
    lock.releaseLock();
  }
}

function trackOrder(query) {
  try {
    var orderId = sanitizeStr_(query && query.orderId, 40).toUpperCase();
    var phone = normalizePhone_(query && query.phone);
    if (!orderId && !phone) return fail_('Masukkan Order ID atau Nomor Telepon.');

    var orders = readAll_('Orders');
    var matched = orders.filter(function(o) {
      if (orderId && String(o.order_id).toUpperCase() === orderId) return true;
      if (phone && normalizePhone_(o.customer_phone) === phone) return true;
      return false;
    }).sort(function(a, b) {
      return String(b.created_at).localeCompare(String(a.created_at));
    }).slice(0, 20);

    if (matched.length === 0) return ok_([], 'Pesanan tidak ditemukan.');

    var orderIds = matched.map(function(m) { return m.order_id; });
    var allItems = readAll_('OrderItems').filter(function(it) { return orderIds.indexOf(it.order_id) !== -1; });
    var allHist = readAll_('OrderStatusHistory').filter(function(h) { return orderIds.indexOf(h.order_id) !== -1; });

    var result = matched.map(function(o) {
      return {
        order_id: o.order_id,
        source: o.source,
        order_type: o.order_type,
        customer_name: o.customer_name,
        customer_phone: o.customer_phone,
        address: o.address,
        subtotal: toNum_(o.subtotal),
        discount_amount: toNum_(o.discount_amount),
        tax_amount: toNum_(o.tax_amount),
        shipping_fee: toNum_(o.shipping_fee),
        total: toNum_(o.total),
        payment_method: o.payment_method,
        status: o.status,
        is_scheduled: toBool_(o.is_scheduled),
        scheduled_time: o.scheduled_time,
        created_at: o.created_at,
        items: allItems.filter(function(it) { return it.order_id === o.order_id; }),
        history: allHist.filter(function(h) { return h.order_id === o.order_id; })
      };
    });

    return ok_(result);
  } catch (err) {
    return fail_(err.message);
  }
}

/* ============================================================================
 * DASHBOARD & POS KASIR
 * ============================================================================ */

function getDashboardData(token) {
  try {
    requireAuth_(token, ['OWNER', 'ADMIN', 'KASIR']);
    var today = todayDate_();
    var orders = readAll_('Orders');
    var trxs = readAll_('Transactions');
    var trxItems = readAll_('TransactionItems');
    var expenses = readAll_('Expenses');

    var todayTrxs = trxs.filter(function(t) { return String(t.created_at).slice(0, 10) === today && t.status !== 'Dibatalkan'; });
    var omsetHariIni = 0;
    var hppHariIni = 0;
    todayTrxs.forEach(function(t) {
      omsetHariIni += toNum_(t.total, 0);
      hppHariIni += toNum_(t.hpp_total, 0);
    });

    var todayExpenses = 0;
    expenses.forEach(function(e) {
      if (String(e.date).slice(0, 10) === today) todayExpenses += toNum_(e.amount, 0);
    });

    var pendingCount = orders.filter(function(o) { return o.status === 'Menunggu Konfirmasi'; }).length;
    var processingCount = orders.filter(function(o) {
      return ['Dikonfirmasi', 'Sedang Diproses', 'Sedang Dimasak', 'Siap Diambil / Dalam Pengiriman'].indexOf(o.status) !== -1;
    }).length;

    // Produk Terlaris & Cek HPP
    var prodStats = {};
    var hasMissingHpp = false;
    trxItems.forEach(function(it) {
      if (toNum_(it.hpp_snapshot, 0) <= 0) hasMissingHpp = true;
      var key = it.product_name_snapshot;
      if (!prodStats[key]) prodStats[key] = { name: key, qty: 0, revenue: 0 };
      prodStats[key].qty += toNum_(it.qty, 0);
      prodStats[key].revenue += toNum_(it.subtotal, 0);
    });
    var topProducts = Object.keys(prodStats).map(function(k) { return prodStats[k]; })
      .sort(function(a, b) { return b.qty - a.qty; }).slice(0, 5);

    // Chart 7 hari terakhir
    var chartLabels = [];
    var chartValues = [];
    var now = new Date();
    for (var d = 6; d >= 0; d--) {
      var dt = new Date(now.getTime() - d * 86400000);
      var ds = Utilities.formatDate(dt, TIMEZONE, 'yyyy-MM-dd');
      chartLabels.push(Utilities.formatDate(dt, TIMEZONE, 'dd/MM'));
      var daySum = 0;
      trxs.forEach(function(t) {
        if (String(t.created_at).slice(0, 10) === ds && t.status !== 'Dibatalkan') daySum += toNum_(t.total, 0);
      });
      chartValues.push(daySum);
    }

    var recentTrxs = trxs.slice().sort(function(a, b) {
      return String(b.created_at).localeCompare(String(a.created_at));
    }).slice(0, 8);

    return ok_({
      omsetHariIni: omsetHariIni,
      jumlahTransaksi: todayTrxs.length,
      pesananMenunggu: pendingCount,
      pesananDiproses: processingCount,
      totalPengeluaranHariIni: todayExpenses,
      labaKotorHariIni: omsetHariIni - hppHariIni,
      estimasiLabaBersihHariIni: omsetHariIni - hppHariIni - todayExpenses,
      hasMissingHpp: hasMissingHpp,
      topProducts: topProducts,
      chart: { labels: chartLabels, values: chartValues },
      recentTransactions: recentTrxs
    });
  } catch (err) {
    return fail_(err.message);
  }
}

function createPosTransaction(token, payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var user = requireAuth_(token, ['OWNER', 'ADMIN', 'KASIR']);
    var paymentMethod = sanitizeStr_(payload.paymentMethod, 40);
    if (!paymentMethod) return fail_('Metode pembayaran wajib dipilih.');
    if (paymentMethod.toUpperCase().indexOf('COD') !== -1) {
      return fail_('Sistem tidak menerima metode pembayaran COD.');
    }

    var calc = buildValidatedOrderCalc_(payload, false);
    var paidAmount = toNum_(payload.paidAmount, 0);
    if (paidAmount < calc.total) {
      return fail_('Uang diterima (Rp ' + paidAmount.toLocaleString('id-ID') + ') kurang dari total tagihan (Rp ' + calc.total.toLocaleString('id-ID') + ').');
    }
    var changeAmount = paidAmount - calc.total;

    var custName = sanitizeStr_(payload.customerName, 100) || 'Pelanggan Walk-In';
    var custPhone = normalizePhone_(payload.customerPhone);
    var orderType = sanitizeStr_(payload.orderType || 'DINE_IN', 30).toUpperCase();
    var now = nowIso_();
    var orderId = generateUniqueOrderId_();
    var trxId = 'TRX-' + Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd') + '-' + Math.floor(1000 + Math.random() * 9000);

    var customerId = '';
    if (custPhone) {
      customerId = upsertCustomer_(custName, custPhone, '', '', '', calc.total, now);
    }

    // Kurangi stok
    var allProds = readAll_('Products');
    calc.items.forEach(function(item) {
      var p = allProds.filter(function(pr) { return pr.id === item.productId; })[0];
      if (p) {
        p.stock = Math.max(0, toNum_(p.stock, 0) - item.qty);
        if (p.stock === 0) p.is_available = false;
        p.updated_at = now;
        updateRowByIndex_('Products', p._rowIndex, p);
      }
    });

    var initStatus = sanitizeStr_(payload.initialStatus, 40) || 'Selesai';
    if (ORDER_STATUSES.indexOf(initStatus) === -1) initStatus = 'Selesai';

    appendRows_('Orders', [{
      order_id: orderId,
      source: 'POS',
      order_type: orderType,
      customer_id: customerId,
      customer_name: custName,
      customer_phone: custPhone,
      address: sanitizeStr_(payload.address, 300),
      delivery_note: sanitizeStr_(payload.note, 200),
      lat: '',
      lng: '',
      distance_km: 0,
      distance_method: 'POS',
      subtotal: calc.subtotal,
      discount_amount: calc.discountAmount,
      promo_code: calc.promoCode,
      tax_amount: calc.taxAmount,
      tax_percent: calc.taxPercent,
      shipping_fee: calc.shippingFee,
      total: calc.total,
      payment_method: paymentMethod,
      payment_status: 'PAID',
      paid_amount: paidAmount,
      change_amount: changeAmount,
      status: initStatus,
      is_scheduled: false,
      scheduled_time: '',
      cashier_user_id: user.id,
      created_at: now,
      updated_at: now
    }]);

    appendRows_('OrderItems', calc.items.map(function(it) {
      return {
        id: genId_('OIT'),
        order_id: orderId,
        product_id: it.productId,
        product_name_snapshot: it.productName,
        variant_id: it.variantId,
        variant_name_snapshot: it.variantName,
        unit_price_snapshot: it.unitPrice,
        hpp_snapshot: it.hpp,
        qty: it.qty,
        subtotal: it.subtotal,
        note: it.note
      };
    }));

    appendRows_('OrderStatusHistory', [{
      id: genId_('OSH'),
      order_id: orderId,
      status: initStatus,
      note: 'Transaksi POS oleh kasir ' + user.username,
      changed_by: user.username,
      changed_at: now
    }]);

    appendRows_('Transactions', [{
      trx_id: trxId,
      order_id: orderId,
      source: 'POS',
      order_type: orderType,
      customer_id: customerId,
      customer_name: custName,
      customer_phone: custPhone,
      subtotal: calc.subtotal,
      discount_amount: calc.discountAmount,
      tax_amount: calc.taxAmount,
      shipping_fee: calc.shippingFee,
      total: calc.total,
      hpp_total: calc.hppTotal,
      payment_method: paymentMethod,
      paid_amount: paidAmount,
      change_amount: changeAmount,
      cashier_id: user.id,
      cashier_name: user.fullName || user.username,
      status: 'PAID',
      created_at: now
    }]);

    appendRows_('TransactionItems', calc.items.map(function(it) {
      return {
        id: genId_('TIT'),
        trx_id: trxId,
        order_id: orderId,
        product_id: it.productId,
        product_name_snapshot: it.productName,
        variant_name_snapshot: it.variantName,
        unit_price_snapshot: it.unitPrice,
        hpp_snapshot: it.hpp,
        qty: it.qty,
        subtotal: it.subtotal,
        note: it.note
      };
    }));

    writeAudit_(user, 'CREATE_POS_TRANSACTION', 'ORDER', orderId, null, { trx_id: trxId, total: calc.total, payment_method: paymentMethod });
    clearPublicCache_();

    return getReceiptData(token, orderId);
  } catch (err) {
    return fail_(err.message);
  } finally {
    lock.releaseLock();
  }
}

/* ============================================================================
 * PESANAN MASUK & STATUS + STRUK/INVOICE
 * ============================================================================ */

function getOrdersAdmin(token, filters) {
  try {
    requireAuth_(token, ['OWNER', 'ADMIN', 'KASIR']);
    filters = filters || {};
    var statusFilter = sanitizeStr_(filters.status, 50);
    var sourceFilter = sanitizeStr_(filters.source, 20);
    var search = sanitizeStr_(filters.search, 80).toLowerCase();
    var page = Math.max(1, toNum_(filters.page, 1));
    var pageSize = Math.min(100, Math.max(5, toNum_(filters.pageSize, 20)));

    var orders = readAll_('Orders').sort(function(a, b) {
      return String(b.created_at).localeCompare(String(a.created_at));
    });

    var filtered = orders.filter(function(o) {
      if (statusFilter && statusFilter !== 'ALL' && o.status !== statusFilter) return false;
      if (sourceFilter && sourceFilter !== 'ALL' && o.source !== sourceFilter) return false;
      if (search) {
        var hay = (o.order_id + ' ' + o.customer_name + ' ' + o.customer_phone).toLowerCase();
        if (hay.indexOf(search) === -1) return false;
      }
      return true;
    });

    var totalItems = filtered.length;
    var slice = filtered.slice((page - 1) * pageSize, page * pageSize);
    var ids = slice.map(function(o) { return o.order_id; });
    var items = readAll_('OrderItems').filter(function(it) { return ids.indexOf(it.order_id) !== -1; });
    var hist = readAll_('OrderStatusHistory').filter(function(h) { return ids.indexOf(h.order_id) !== -1; });

    var rows = slice.map(function(o) {
      o.items = items.filter(function(it) { return it.order_id === o.order_id; });
      o.history = hist.filter(function(h) { return h.order_id === o.order_id; });
      return o;
    });

    return ok_({
      rows: rows,
      total: totalItems,
      page: page,
      pageSize: pageSize,
      totalPages: Math.max(1, Math.ceil(totalItems / pageSize))
    });
  } catch (err) {
    return fail_(err.message);
  }
}

function updateOrderStatus(token, orderId, newStatus, note) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var user = requireAuth_(token, ['OWNER', 'ADMIN', 'KASIR']);
    if (ORDER_STATUSES.indexOf(newStatus) === -1) return fail_('Status pesanan tidak valid.');

    var orders = readAll_('Orders');
    var ord = orders.filter(function(o) { return o.order_id === orderId; })[0];
    if (!ord) return fail_('Pesanan tidak ditemukan.');

    var oldStatus = ord.status;
    var now = nowIso_();
    ord.status = newStatus;
    ord.updated_at = now;

    // Jika dikonfirmasi / dibayar, catat ke Transactions jika belum ada (khusus Online)
    var trxs = readAll_('Transactions');
    var existingTrx = trxs.filter(function(t) { return t.order_id === orderId; })[0];

    if (newStatus === 'Dikonfirmasi' || newStatus === 'Selesai') {
      ord.payment_status = 'PAID';
      if (toNum_(ord.paid_amount, 0) === 0) ord.paid_amount = ord.total;

      if (!existingTrx) {
        var orderItems = readAll_('OrderItems').filter(function(it) { return it.order_id === orderId; });
        var hppSum = 0;
        orderItems.forEach(function(it) { hppSum += toNum_(it.hpp_snapshot, 0) * toNum_(it.qty, 0); });
        var trxId = 'TRX-' + Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd') + '-' + Math.floor(1000 + Math.random() * 9000);
        appendRows_('Transactions', [{
          trx_id: trxId,
          order_id: orderId,
          source: ord.source,
          order_type: ord.order_type,
          customer_id: ord.customer_id,
          customer_name: ord.customer_name,
          customer_phone: ord.customer_phone,
          subtotal: ord.subtotal,
          discount_amount: ord.discount_amount,
          tax_amount: ord.tax_amount,
          shipping_fee: ord.shipping_fee,
          total: ord.total,
          hpp_total: hppSum,
          payment_method: ord.payment_method,
          paid_amount: ord.total,
          change_amount: 0,
          cashier_id: user.id,
          cashier_name: user.fullName || user.username,
          status: 'PAID',
          created_at: now
        }]);
        appendRows_('TransactionItems', orderItems.map(function(it) {
          return {
            id: genId_('TIT'),
            trx_id: trxId,
            order_id: orderId,
            product_id: it.product_id,
            product_name_snapshot: it.product_name_snapshot,
            variant_name_snapshot: it.variant_name_snapshot,
            unit_price_snapshot: it.unit_price_snapshot,
            hpp_snapshot: it.hpp_snapshot,
            qty: it.qty,
            subtotal: it.subtotal,
            note: it.note
          };
        }));
      }
    }

    if (newStatus === 'Dibatalkan' && existingTrx) {
      existingTrx.status = 'Dibatalkan';
      updateRowByIndex_('Transactions', existingTrx._rowIndex, existingTrx);
    }

    // Kembalikan stok jika baru dibatalkan
    if (newStatus === 'Dibatalkan' && oldStatus !== 'Dibatalkan') {
      var ordItems = readAll_('OrderItems').filter(function(it) { return it.order_id === orderId; });
      var allProds = readAll_('Products');
      ordItems.forEach(function(it) {
        var p = allProds.filter(function(pr) { return pr.id === it.product_id; })[0];
        if (p) {
          p.stock = toNum_(p.stock, 0) + toNum_(it.qty, 0);
          p.is_available = true;
          p.updated_at = now;
          updateRowByIndex_('Products', p._rowIndex, p);
        }
      });
      clearPublicCache_();
    }

    updateRowByIndex_('Orders', ord._rowIndex, ord);
    appendRows_('OrderStatusHistory', [{
      id: genId_('OSH'),
      order_id: orderId,
      status: newStatus,
      note: sanitizeStr_(note, 200) || ('Status diubah ke ' + newStatus),
      changed_by: user.username,
      changed_at: now
    }]);

    writeAudit_(user, newStatus === 'Dibatalkan' ? 'CANCEL_ORDER' : 'UPDATE_ORDER_STATUS', 'ORDER', orderId, { status: oldStatus }, { status: newStatus, note: note || '' });
    return ok_({ orderId: orderId, status: newStatus }, 'Status pesanan diperbarui.');
  } catch (err) {
    return fail_(err.message);
  } finally {
    lock.releaseLock();
  }
}

function getReceiptData(token, orderId) {
  try {
    requireAuth_(token, ['OWNER', 'ADMIN', 'KASIR']);
    var settingsMap = getSettingsMap_();
    var ord = readAll_('Orders').filter(function(o) { return o.order_id === orderId; })[0];
    if (!ord) return fail_('Pesanan tidak ditemukan.');
    var items = readAll_('OrderItems').filter(function(it) { return it.order_id === orderId; });
    return ok_({
      order: ord,
      items: items,
      receiptConfig: {
        restaurant_name: settingsMap.restaurant_name || 'ASAP PEDAS X TEKO',
        address: settingsMap.address || '',
        phone: settingsMap.whatsapp_number || '',
        footer: settingsMap.receipt_footer || 'Terima kasih!',
        paper_size: settingsMap.receipt_paper_size || '80mm'
      }
    });
  } catch (err) {
    return fail_(err.message);
  }
}

/* ============================================================================
 * MENU, KATEGORI, BANNER & PROMO CRUD
 * ============================================================================ */

function getProductsAdmin(token) {
  try {
    requireAuth_(token, ['OWNER', 'ADMIN', 'KASIR']);
    var cats = readAll_('Categories').sort(function(a, b) { return toNum_(a.sort_order) - toNum_(b.sort_order); });
    var vars = readAll_('Variants');
    var prods = readAll_('Products').filter(function(p) { return !toBool_(p.is_deleted); }).map(function(p) {
      p.variants = vars.filter(function(v) { return v.product_id === p.id; });
      return p;
    });
    return ok_({ products: prods, categories: cats });
  } catch (err) {
    return fail_(err.message);
  }
}

function saveProduct(token, payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var user = requireAuth_(token, ['OWNER', 'ADMIN']);
    var name = sanitizeStr_(payload.name, 120);
    var catId = sanitizeStr_(payload.category_id, 40);
    var price = toNum_(payload.price, -1);
    if (!name || !catId || price < 0) return fail_('Nama, kategori, dan harga valid wajib diisi.');

    var now = nowIso_();
    var prods = readAll_('Products');
    var prodId = sanitizeStr_(payload.id, 40);
    var existing = prodId ? prods.filter(function(p) { return p.id === prodId; })[0] : null;
    var before = existing ? JSON.parse(JSON.stringify(existing)) : null;

    var record = {
      id: existing ? existing.id : genId_('PRD'),
      category_id: catId,
      name: name,
      description: sanitizeStr_(payload.description, 500),
      image_url: sanitizeStr_(payload.image_url, 500),
      price: price,
      promo_price: payload.promo_price !== '' && payload.promo_price !== null ? toNum_(payload.promo_price, '') : '',
      hpp: payload.hpp !== '' && payload.hpp !== null ? toNum_(payload.hpp, 0) : 0,
      stock: Math.max(0, parseInt(payload.stock || 0, 10)),
      is_available: toBool_(payload.is_available),
      is_popular: toBool_(payload.is_popular),
      is_deleted: false,
      created_at: existing ? existing.created_at : now,
      updated_at: now
    };

    if (existing) {
      updateRowByIndex_('Products', existing._rowIndex, record);
      var actionName = toNum_(before.price) !== record.price ? 'UPDATE_PRODUCT_PRICE' : 'UPDATE_PRODUCT';
      writeAudit_(user, actionName, 'PRODUCT', record.id, before, record);
    } else {
      appendRows_('Products', [record]);
      writeAudit_(user, 'CREATE_PRODUCT', 'PRODUCT', record.id, null, record);
    }

    // Update Variants
    var existingVars = readAll_('Variants');
    for (var i = existingVars.length - 1; i >= 0; i--) {
      if (existingVars[i].product_id === record.id) {
        deleteRowByIndex_('Variants', existingVars[i]._rowIndex);
      }
    }
    if (payload.variants && payload.variants.length > 0) {
      var varRows = payload.variants.filter(function(v) { return sanitizeStr_(v.name, 80); }).map(function(v) {
        return {
          id: sanitizeStr_(v.id, 40) || genId_('VAR'),
          product_id: record.id,
          name: sanitizeStr_(v.name, 80),
          price_adjustment: toNum_(v.price_adjustment, 0),
          is_active: v.is_active !== undefined ? toBool_(v.is_active) : true
        };
      });
      appendRows_('Variants', varRows);
    }

    clearPublicCache_();
    return ok_(record, 'Menu berhasil disimpan.');
  } catch (err) {
    return fail_(err.message);
  } finally {
    lock.releaseLock();
  }
}

function deleteProduct(token, productId) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var user = requireAuth_(token, ['OWNER', 'ADMIN']);
    var prods = readAll_('Products');
    var p = prods.filter(function(x) { return x.id === productId; })[0];
    if (!p) return fail_('Produk tidak ditemukan.');
    var before = JSON.parse(JSON.stringify(p));
    p.is_deleted = true;
    p.is_available = false;
    p.updated_at = nowIso_();
    updateRowByIndex_('Products', p._rowIndex, p);
    writeAudit_(user, 'DELETE_PRODUCT', 'PRODUCT', productId, before, { is_deleted: true });
    clearPublicCache_();
    return ok_(null, 'Menu berhasil dihapus (riwayat transaksi tetap aman).');
  } catch (err) {
    return fail_(err.message);
  } finally {
    lock.releaseLock();
  }
}

function saveCategory(token, payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var user = requireAuth_(token, ['OWNER', 'ADMIN']);
    var name = sanitizeStr_(payload.name, 80);
    if (!name) return fail_('Nama kategori wajib diisi.');
    var now = nowIso_();
    var cats = readAll_('Categories');
    var existing = payload.id ? cats.filter(function(c) { return c.id === payload.id; })[0] : null;
    var rec = {
      id: existing ? existing.id : genId_('CAT'),
      name: name,
      slug: name.toLowerCase().replace(/[^a-z0-9]+/g, '-'),
      sort_order: toNum_(payload.sort_order, cats.length + 1),
      is_active: payload.is_active !== undefined ? toBool_(payload.is_active) : true,
      created_at: existing ? existing.created_at : now,
      updated_at: now
    };
    if (existing) {
      updateRowByIndex_('Categories', existing._rowIndex, rec);
      writeAudit_(user, 'UPDATE_CATEGORY', 'CATEGORY', rec.id, existing, rec);
    } else {
      appendRows_('Categories', [rec]);
      writeAudit_(user, 'CREATE_CATEGORY', 'CATEGORY', rec.id, null, rec);
    }
    clearPublicCache_();
    return ok_(rec, 'Kategori berhasil disimpan.');
  } catch (err) {
    return fail_(err.message);
  } finally {
    lock.releaseLock();
  }
}

function deleteCategory(token, catId) {
  try {
    var user = requireAuth_(token, ['OWNER', 'ADMIN']);
    var prods = readAll_('Products').filter(function(p) { return !toBool_(p.is_deleted) && p.category_id === catId; });
    if (prods.length > 0) return fail_('Kategori masih digunakan oleh ' + prods.length + ' menu aktif. Nonaktifkan saja atau pindahkan menu.');
    var cats = readAll_('Categories');
    var c = cats.filter(function(x) { return x.id === catId; })[0];
    if (!c) return fail_('Kategori tidak ditemukan.');
    deleteRowByIndex_('Categories', c._rowIndex);
    writeAudit_(user, 'DELETE_CATEGORY', 'CATEGORY', catId, c, null);
    clearPublicCache_();
    return ok_(null, 'Kategori dihapus.');
  } catch (err) {
    return fail_(err.message);
  }
}

function getBannersAdmin(token) {
  try {
    requireAuth_(token, ['OWNER', 'ADMIN']);
    return ok_(readAll_('Banners').sort(function(a, b) { return toNum_(a.sort_order) - toNum_(b.sort_order); }));
  } catch (err) {
    return fail_(err.message);
  }
}

function saveBanner(token, payload) {
  try {
    var user = requireAuth_(token, ['OWNER', 'ADMIN']);
    var now = nowIso_();
    var list = readAll_('Banners');
    var ex = payload.id ? list.filter(function(b) { return b.id === payload.id; })[0] : null;
    var rec = {
      id: ex ? ex.id : genId_('BNR'),
      image_url: sanitizeStr_(payload.image_url, 500),
      headline: sanitizeStr_(payload.headline, 150),
      subheadline: sanitizeStr_(payload.subheadline, 250),
      button_text: sanitizeStr_(payload.button_text, 50) || 'Lihat Menu',
      button_link: sanitizeStr_(payload.button_link, 200) || '#menu',
      sort_order: toNum_(payload.sort_order, list.length + 1),
      is_active: payload.is_active !== undefined ? toBool_(payload.is_active) : true,
      created_at: ex ? ex.created_at : now
    };
    if (ex) updateRowByIndex_('Banners', ex._rowIndex, rec);
    else appendRows_('Banners', [rec]);
    writeAudit_(user, ex ? 'UPDATE_BANNER' : 'CREATE_BANNER', 'BANNER', rec.id, ex, rec);
    clearPublicCache_();
    return ok_(rec, 'Banner disimpan.');
  } catch (err) {
    return fail_(err.message);
  }
}

function deleteBanner(token, id) {
  try {
    var user = requireAuth_(token, ['OWNER', 'ADMIN']);
    var ex = readAll_('Banners').filter(function(b) { return b.id === id; })[0];
    if (!ex) return fail_('Banner tidak ditemukan.');
    deleteRowByIndex_('Banners', ex._rowIndex);
    writeAudit_(user, 'DELETE_BANNER', 'BANNER', id, ex, null);
    clearPublicCache_();
    return ok_(null, 'Banner dihapus.');
  } catch (err) {
    return fail_(err.message);
  }
}

function getPromotionsAdmin(token) {
  try {
    requireAuth_(token, ['OWNER', 'ADMIN']);
    return ok_(readAll_('Promotions'));
  } catch (err) {
    return fail_(err.message);
  }
}

function savePromotion(token, payload) {
  try {
    var user = requireAuth_(token, ['OWNER', 'ADMIN']);
    var code = sanitizeStr_(payload.code, 30).toUpperCase();
    if (!code) return fail_('Kode promo wajib diisi.');
    var list = readAll_('Promotions');
    var ex = payload.id ? list.filter(function(p) { return p.id === payload.id; })[0] : null;
    var rec = {
      id: ex ? ex.id : genId_('PRM'),
      code: code,
      name: sanitizeStr_(payload.name, 100),
      type: payload.type === 'FIXED' ? 'FIXED' : 'PERCENT',
      value: toNum_(payload.value, 0),
      min_order: toNum_(payload.min_order, 0),
      max_discount: toNum_(payload.max_discount, 0),
      start_date: sanitizeStr_(payload.start_date, 20) || todayDate_(),
      end_date: sanitizeStr_(payload.end_date, 20) || '2028-12-31',
      is_active: payload.is_active !== undefined ? toBool_(payload.is_active) : true,
      created_at: ex ? ex.created_at : nowIso_()
    };
    if (ex) updateRowByIndex_('Promotions', ex._rowIndex, rec);
    else appendRows_('Promotions', [rec]);
    clearPublicCache_();
    return ok_(rec, 'Promo disimpan.');
  } catch (err) {
    return fail_(err.message);
  }
}

function deletePromotion(token, id) {
  try {
    requireAuth_(token, ['OWNER', 'ADMIN']);
    var ex = readAll_('Promotions').filter(function(p) { return p.id === id; })[0];
    if (!ex) return fail_('Promo tidak ditemukan.');
    deleteRowByIndex_('Promotions', ex._rowIndex);
    clearPublicCache_();
    return ok_(null, 'Promo dihapus.');
  } catch (err) {
    return fail_(err.message);
  }
}

/* ============================================================================
 * LAPORAN PENJUALAN, PELANGGAN, PENGELUARAN & AUDIT LOG
 * ============================================================================ */

function getDateRange_(preset, startStr, endStr) {
  var now = new Date();
  var fmt = function(d) { return Utilities.formatDate(d, TIMEZONE, 'yyyy-MM-dd'); };
  var today = fmt(now);
  var start = today, end = today;

  if (preset === 'yesterday') {
    var y = new Date(now.getTime() - 86400000);
    start = fmt(y); end = fmt(y);
  } else if (preset === 'this_week') {
    var day = Number(Utilities.formatDate(now, TIMEZONE, 'u'));
    start = fmt(new Date(now.getTime() - (day - 1) * 86400000));
  } else if (preset === 'last_week') {
    var d2 = Number(Utilities.formatDate(now, TIMEZONE, 'u'));
    var thisMon = new Date(now.getTime() - (d2 - 1) * 86400000);
    start = fmt(new Date(thisMon.getTime() - 7 * 86400000));
    end = fmt(new Date(thisMon.getTime() - 86400000));
  } else if (preset === 'this_month') {
    start = today.slice(0, 8) + '01';
  } else if (preset === 'last_month') {
    var firstThis = new Date(now.getFullYear(), now.getMonth(), 1);
    var lastPrev = new Date(firstThis.getTime() - 86400000);
    var firstPrev = new Date(lastPrev.getFullYear(), lastPrev.getMonth(), 1);
    start = fmt(firstPrev); end = fmt(lastPrev);
  } else if (preset === 'ytd') {
    start = today.slice(0, 4) + '-01-01';
  } else if (preset === 'custom' && startStr && endStr) {
    start = startStr; end = endStr;
  }
  return { start: start, end: end };
}

function getSalesReport(token, filterPayload) {
  try {
    requireAuth_(token, ['OWNER', 'ADMIN']);
    filterPayload = filterPayload || {};
    var range = getDateRange_(filterPayload.preset || 'today', filterPayload.startDate, filterPayload.endDate);

    var trxs = readAll_('Transactions').filter(function(t) {
      var d = String(t.created_at).slice(0, 10);
      return t.status !== 'Dibatalkan' && d >= range.start && d <= range.end;
    });
    var trxIds = trxs.map(function(t) { return t.trx_id; });
    var items = readAll_('TransactionItems').filter(function(it) { return trxIds.indexOf(it.trx_id) !== -1; });
    var expenses = readAll_('Expenses').filter(function(e) {
      var d = String(e.date).slice(0, 10);
      return d >= range.start && d <= range.end;
    });

    var totalOmset = 0, totalHpp = 0, totalExpenses = 0;
    var byDay = {};
    trxs.forEach(function(t) {
      totalOmset += toNum_(t.total);
      totalHpp += toNum_(t.hpp_total);
      var d = String(t.created_at).slice(0, 10);
      byDay[d] = (byDay[d] || 0) + toNum_(t.total);
    });
    expenses.forEach(function(e) { totalExpenses += toNum_(e.amount); });

    var prodMap = {};
    var hasMissingHpp = false;
    items.forEach(function(it) {
      if (toNum_(it.hpp_snapshot, 0) <= 0) hasMissingHpp = true;
      var k = it.product_name_snapshot;
      if (!prodMap[k]) prodMap[k] = { name: k, qty: 0, revenue: 0 };
      prodMap[k].qty += toNum_(it.qty);
      prodMap[k].revenue += toNum_(it.subtotal);
    });
    var topProducts = Object.keys(prodMap).map(function(k) { return prodMap[k]; })
      .sort(function(a, b) { return b.qty - a.qty; }).slice(0, 10);

    var dayKeys = Object.keys(byDay).sort();
    return ok_({
      range: range,
      summary: {
        totalOmset: totalOmset,
        trxCount: trxs.length,
        avgTrx: trxs.length > 0 ? Math.round(totalOmset / trxs.length) : 0,
        totalHpp: totalHpp,
        totalExpenses: totalExpenses,
        grossProfit: totalOmset - totalHpp,
        estimatedNetProfit: totalOmset - totalHpp - totalExpenses,
        hasMissingHpp: hasMissingHpp
      },
      topProducts: topProducts,
      chart: { labels: dayKeys, values: dayKeys.map(function(k) { return byDay[k]; }) },
      transactions: trxs.sort(function(a, b) { return String(b.created_at).localeCompare(String(a.created_at)); })
    });
  } catch (err) {
    return fail_(err.message);
  }
}

function getCustomersAdmin(token, filters) {
  try {
    requireAuth_(token, ['OWNER', 'ADMIN', 'KASIR']);
    filters = filters || {};
    var q = sanitizeStr_(filters.search, 80).toLowerCase();
    var customers = readAll_('Customers').filter(function(c) {
      if (!q) return true;
      return (c.name + ' ' + c.phone + ' ' + c.last_address).toLowerCase().indexOf(q) !== -1;
    }).sort(function(a, b) { return toNum_(b.total_spent) - toNum_(a.total_spent); });

    var orders = readAll_('Orders');
    var result = customers.map(function(c) {
      c.orders = orders.filter(function(o) {
        return o.customer_id === c.id || normalizePhone_(o.customer_phone) === normalizePhone_(c.phone);
      }).sort(function(a, b) { return String(b.created_at).localeCompare(String(a.created_at)); }).slice(0, 15);
      return c;
    });
    return ok_(result);
  } catch (err) {
    return fail_(err.message);
  }
}

function getExpensesAdmin(token, filters) {
  try {
    requireAuth_(token, ['OWNER', 'ADMIN']);
    var list = readAll_('Expenses').sort(function(a, b) {
      return String(b.date).localeCompare(String(a.date));
    });
    return ok_(list);
  } catch (err) {
    return fail_(err.message);
  }
}

function saveExpense(token, payload) {
  try {
    var user = requireAuth_(token, ['OWNER', 'ADMIN']);
    var amount = toNum_(payload.amount, 0);
    var cat = sanitizeStr_(payload.category, 60);
    if (!cat || amount <= 0) return fail_('Kategori dan nominal pengeluaran (> 0) wajib diisi.');
    var rec = {
      id: genId_('EXP'),
      date: sanitizeStr_(payload.date, 20) || todayDate_(),
      category: cat,
      amount: amount,
      description: sanitizeStr_(payload.description, 250),
      user_id: user.id,
      username: user.username,
      created_at: nowIso_()
    };
    appendRows_('Expenses', [rec]);
    writeAudit_(user, 'CREATE_EXPENSE', 'EXPENSE', rec.id, null, rec);
    return ok_(rec, 'Pengeluaran dicatat.');
  } catch (err) {
    return fail_(err.message);
  }
}

function getAuditLogsAdmin(token, filters) {
  try {
    // Hanya OWNER yang berhak membaca AuditLog (read-only, tanpa fungsi edit/hapus)
    requireAuth_(token, ['OWNER']);
    filters = filters || {};
    var page = Math.max(1, toNum_(filters.page, 1));
    var pageSize = Math.min(100, Math.max(10, toNum_(filters.pageSize, 30)));
    var search = sanitizeStr_(filters.search, 60).toLowerCase();

    var logs = readAll_('AuditLogs').sort(function(a, b) {
      return String(b.created_at).localeCompare(String(a.created_at));
    }).filter(function(l) {
      if (!search) return true;
      return (l.username + ' ' + l.action + ' ' + l.target_type + ' ' + l.target_id).toLowerCase().indexOf(search) !== -1;
    });

    var total = logs.length;
    return ok_({
      rows: logs.slice((page - 1) * pageSize, page * pageSize),
      total: total,
      page: page,
      totalPages: Math.max(1, Math.ceil(total / pageSize))
    });
  } catch (err) {
    return fail_(err.message);
  }
}

/* ============================================================================
 * PENGATURAN, LIBUR, SOSIAL MEDIA & MANAJEMEN USER
 * ============================================================================ */

function getSettingsAdmin(token) {
  try {
    requireAuth_(token, ['OWNER', 'ADMIN']);
    var settingsMap = getSettingsMap_();
    return ok_({
      settings: settingsMap,
      holidays: readAll_('Holidays').sort(function(a, b) { return String(b.date).localeCompare(String(a.date)); }),
      socialMedia: readAll_('SocialMedia').sort(function(a, b) { return toNum_(a.sort_order) - toNum_(b.sort_order); })
    });
  } catch (err) {
    return fail_(err.message);
  }
}

function saveSettingsAdmin(token, payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var user = requireAuth_(token, ['OWNER', 'ADMIN']);
    var existingRows = readAll_('Settings');
    var beforeMap = {};
    var rowByKey = {};
    existingRows.forEach(function(r) {
      beforeMap[r.key] = r.value;
      rowByKey[r.key] = r;
    });

    var now = nowIso_();
    var toAppend = [];
    Object.keys(payload).forEach(function(k) {
      var val = typeof payload[k] === 'object' ? JSON.stringify(payload[k]) : String(payload[k]);
      if (rowByKey[k]) {
        var r = rowByKey[k];
        r.value = val;
        r.updated_at = now;
        updateRowByIndex_('Settings', r._rowIndex, r);
      } else {
        toAppend.push({ key: k, value: val, group: 'custom', description: k, updated_at: now });
      }
    });
    if (toAppend.length > 0) appendRows_('Settings', toAppend);

    writeAudit_(user, 'UPDATE_SETTINGS', 'SETTINGS', 'GLOBAL', beforeMap, payload);
    clearPublicCache_();
    return ok_(null, 'Pengaturan berhasil disimpan.');
  } catch (err) {
    return fail_(err.message);
  } finally {
    lock.releaseLock();
  }
}

function saveHoliday(token, payload) {
  try {
    var user = requireAuth_(token, ['OWNER', 'ADMIN']);
    var date = sanitizeStr_(payload.date, 20);
    var name = sanitizeStr_(payload.name, 100);
    if (!date || !name) return fail_('Tanggal dan nama hari libur wajib diisi.');
    var rec = {
      id: genId_('HOL'),
      date: date,
      name: name,
      is_closed: payload.is_closed !== undefined ? toBool_(payload.is_closed) : true,
      note: sanitizeStr_(payload.note, 200),
      created_at: nowIso_()
    };
    appendRows_('Holidays', [rec]);
    writeAudit_(user, 'CREATE_HOLIDAY', 'HOLIDAY', rec.id, null, rec);
    clearPublicCache_();
    return ok_(rec, 'Hari libur ditambahkan.');
  } catch (err) {
    return fail_(err.message);
  }
}

function deleteHoliday(token, id) {
  try {
    var user = requireAuth_(token, ['OWNER', 'ADMIN']);
    var ex = readAll_('Holidays').filter(function(h) { return h.id === id; })[0];
    if (!ex) return fail_('Data tidak ditemukan.');
    deleteRowByIndex_('Holidays', ex._rowIndex);
    clearPublicCache_();
    return ok_(null, 'Hari libur dihapus.');
  } catch (err) {
    return fail_(err.message);
  }
}

function saveSocialMedia(token, payload) {
  try {
    var user = requireAuth_(token, ['OWNER', 'ADMIN']);
    var list = readAll_('SocialMedia');
    var ex = payload.id ? list.filter(function(s) { return s.id === payload.id; })[0] : null;
    var rec = {
      id: ex ? ex.id : genId_('SOC'),
      platform: sanitizeStr_(payload.platform, 50),
      label: sanitizeStr_(payload.label, 80),
      url: sanitizeStr_(payload.url, 300),
      icon: sanitizeStr_(payload.icon, 50) || 'bi-link-45deg',
      type: payload.type === 'DELIVERY_APP' ? 'DELIVERY_APP' : 'SOCIAL',
      sort_order: toNum_(payload.sort_order, list.length + 1),
      is_active: payload.is_active !== undefined ? toBool_(payload.is_active) : true
    };
    if (ex) updateRowByIndex_('SocialMedia', ex._rowIndex, rec);
    else appendRows_('SocialMedia', [rec]);
    clearPublicCache_();
    return ok_(rec, 'Tautan media sosial disimpan.');
  } catch (err) {
    return fail_(err.message);
  }
}

function deleteSocialMedia(token, id) {
  try {
    requireAuth_(token, ['OWNER', 'ADMIN']);
    var ex = readAll_('SocialMedia').filter(function(s) { return s.id === id; })[0];
    if (!ex) return fail_('Data tidak ditemukan.');
    deleteRowByIndex_('SocialMedia', ex._rowIndex);
    clearPublicCache_();
    return ok_(null, 'Tautan dihapus.');
  } catch (err) {
    return fail_(err.message);
  }
}

function getUsersAdmin(token) {
  try {
    requireAuth_(token, ['OWNER', 'ADMIN']);
    // Jangan pernah mengirim password_hash atau salt ke frontend
    var users = readAll_('Users').map(function(u) {
      return {
        id: u.id,
        username: u.username,
        full_name: u.full_name,
        role: u.role,
        is_active: toBool_(u.is_active),
        created_at: u.created_at,
        last_login: u.last_login
      };
    });
    return ok_(users);
  } catch (err) {
    return fail_(err.message);
  }
}

function saveUserAdmin(token, payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var actor = requireAuth_(token, ['OWNER']);
    var uname = sanitizeStr_(payload.username, 50).toLowerCase();
    var fullName = sanitizeStr_(payload.full_name, 100);
    var role = ['OWNER', 'ADMIN', 'KASIR'].indexOf(payload.role) !== -1 ? payload.role : 'KASIR';
    if (!uname || !fullName) return fail_('Username dan nama lengkap wajib diisi.');

    var users = readAll_('Users');
    var ex = payload.id ? users.filter(function(u) { return u.id === payload.id; })[0] : null;
    var dup = users.filter(function(u) { return String(u.username).toLowerCase() === uname && (!ex || u.id !== ex.id); })[0];
    if (dup) return fail_('Username sudah digunakan.');

    var now = nowIso_();
    if (ex) {
      var salt = ex.salt;
      var passHash = ex.password_hash;
      if (payload.password && String(payload.password).length >= 6) {
        salt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
        passHash = hashPassword_(String(payload.password), salt);
      }
      var updated = {
        id: ex.id,
        username: uname,
        full_name: fullName,
        role: role,
        password_hash: passHash,
        salt: salt,
        is_active: payload.is_active !== undefined ? toBool_(payload.is_active) : true,
        created_at: ex.created_at,
        updated_at: now,
        last_login: ex.last_login
      };
      updateRowByIndex_('Users', ex._rowIndex, updated);
      writeAudit_(actor, 'UPDATE_USER', 'USER', ex.id, { username: ex.username, role: ex.role }, { username: uname, role: role });
    } else {
      if (!payload.password || String(payload.password).length < 6) {
        return fail_('Password minimal 6 karakter untuk user baru.');
      }
      var newSalt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
      var newRec = {
        id: genId_('USR'),
        username: uname,
        full_name: fullName,
        role: role,
        password_hash: hashPassword_(String(payload.password), newSalt),
        salt: newSalt,
        is_active: true,
        created_at: now,
        updated_at: now,
        last_login: ''
      };
      appendRows_('Users', [newRec]);
      writeAudit_(actor, 'CREATE_USER', 'USER', newRec.id, null, { username: uname, role: role });
    }
    return ok_(null, 'Data user disimpan.');
  } catch (err) {
    return fail_(err.message);
  } finally {
    lock.releaseLock();
  }
}

function deleteUserAdmin(token, userId) {
  try {
    var actor = requireAuth_(token, ['OWNER']);
    if (actor.id === userId) return fail_('Tidak dapat menghapus akun sendiri yang sedang login.');
    var users = readAll_('Users');
    var target = users.filter(function(u) { return u.id === userId; })[0];
    if (!target) return fail_('User tidak ditemukan.');
    deleteRowByIndex_('Users', target._rowIndex);
    writeAudit_(actor, 'DELETE_USER', 'USER', userId, { username: target.username, role: target.role }, null);
    return ok_(null, 'User dihapus.');
  } catch (err) {
    return fail_(err.message);
  }
}

function changeOwnPassword(token, oldPass, newPass) {
  try {
    var actor = requireAuth_(token, ['OWNER', 'ADMIN', 'KASIR']);
    if (!newPass || String(newPass).length < 6) return fail_('Password baru minimal 6 karakter.');
    var users = readAll_('Users');
    var u = users.filter(function(x) { return x.id === actor.id; })[0];
    if (hashPassword_(String(oldPass || ''), String(u.salt)) !== String(u.password_hash)) {
      return fail_('Password lama salah.');
    }
    var newSalt = Utilities.getUuid().replace(/-/g, '').slice(0, 16);
    u.salt = newSalt;
    u.password_hash = hashPassword_(String(newPass), newSalt);
    u.updated_at = nowIso_();
    updateRowByIndex_('Users', u._rowIndex, u);
    writeAudit_(actor, 'CHANGE_OWN_PASSWORD', 'USER', actor.id, null, null);
    return ok_(null, 'Password berhasil diperbarui.');
  } catch (err) {
    return fail_(err.message);
  }
}


</script>
