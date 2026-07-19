// file: code.gs
// ==========================================
// BACKEND ENGINE & CRITICAL CRUD API
// ==========================================

function doGet() {
return HtmlService.createTemplateFromFile('index')
.evaluate()
.setTitle('Web Apps Tracker Keuangan - By Zuyin')
.addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Helper: Bersihkan input string angka rupiah menjadi format float valid murni
function parseFormattedNumber(val) {
if (typeof val === 'number') return val;
if (!val) return 0;
var clean = val.toString().replace(/[^0-9\.,-]/g, "");
if (clean.indexOf(',') !== -1 && clean.indexOf('.') !== -1) {
if (clean.indexOf('.') < clean.indexOf(',')) {
clean = clean.replace(/\./g, "").replace(',', '.');
} else {
clean = clean.replace(/,/g, "");
}
} else if (clean.indexOf('.') !== -1) {
var parts = clean.split('.');
if (parts[parts.length - 1].length === 3) {
clean = clean.replace(/\./g, "");
}
} else if (clean.indexOf(',') !== -1) {
var parts = clean.split(',');
if (parts[parts.length - 1].length === 3) {
clean = clean.replace(/,/g, "");
} else {
clean = clean.replace(',', '.');
}
}
var parsed = parseFloat(clean);
return isNaN(parsed) ? 0 : parsed;
}

// 1. GENERATE DAILY SEQUENTIAL TRANSACTION ID
function generateTransactionId(targetDateStr) {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('transaksi');
var data = sheet.getDataRange().getDisplayValues();

var parts = targetDateStr.split('/');
var yyyymmdd = parts[2] + parts[1] + parts[0];
var prefix = "TX-" + yyyymmdd + "-";

var maxSequence = 0;
for (var i = 1; i < data.length; i++) {
var id = data[i][0];
if (id && id.indexOf(prefix) === 0) {
var seqStr = id.substring(prefix.length);
var seqInt = parseInt(seqStr, 10);
if (!isNaN(seqInt) && seqInt > maxSequence) {
maxSequence = seqInt;
}
}
}

var nextSeq = maxSequence + 1;
var nextSeqStr = ("0000" + nextSeq).slice(-4);
return prefix + nextSeqStr;
}

// 2. GET SYSTEM STATE FOR INITIALIZATION
function getAppData() {
try {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var txSheet = ss.getSheetByName('transaksi');
var txData = txSheet.getDataRange().getDisplayValues();

var categories = getCategories();
var totalPemasukan = 0;
var totalPengeluaran = 0;

for (var i = 1; i < txData.length; i++) {
var tipe = txData[i][3];
var nominal = parseFormattedNumber(txData[i][5]);
if (tipe === 'Pemasukan') {
totalPemasukan += nominal;
} else if (tipe === 'Pengeluaran') {
totalPengeluaran += nominal;
}
}
var walletBalance = totalPemasukan - totalPengeluaran;

return {
walletBalance: walletBalance,
totalPemasukan: totalPemasukan,
totalPengeluaran: totalPengeluaran,
categories: categories
};
} catch(e) {
return { error: e.toString() };
}
}

// 3. SERVER SIDE PAGINATION & BACKEND FILTERING
function getPaginatedTransactions(limit, offset, searchKey, filterType) {
try {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('transaksi');
var data = sheet.getDataRange().getDisplayValues();
var filtered = [];

searchKey = searchKey ? searchKey.toLowerCase().trim() : "";
filterType = filterType ? filterType : "Semua";

for (var i = data.length - 1; i >= 1; i--) {
var row = data[i];
var tanggal = row[1];
var tipe = row[3];
var kategori = row[4];
var nominal = row[5];
var catatan = row[6].toLowerCase();

if (filterType !== "Semua" && tipe !== filterType) continue;

if (searchKey !== "") {
if (kategori.toLowerCase().indexOf(searchKey) === -1 &&
catatan.indexOf(searchKey) === -1 &&
nominal.indexOf(searchKey) === -1 &&
tanggal.indexOf(searchKey) === -1) {
continue;
}
}

filtered.push({
id: row[0],
tanggal: row[1],
jam: row[2],
tipe: row[3],
kategori: row[4],
nominal: parseFormattedNumber(row[5]),
catatan: row[6]
});
}

var chunk = filtered.slice(offset, offset + limit);
return {
transactions: chunk,
hasMore: (offset + limit) < filtered.length
};
} catch(e) {
return { error: e.toString(), transactions: [], hasMore: false };
}
}

// 4. GET METRICS FOR DUAL CHART LABA RUGI REPORT
function getReportData() {
try {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('transaksi');
var data = sheet.getDataRange().getDisplayValues();

var expenseCategoryMap = {};
var incomeCategoryMap = {};
var trendMap = {};

var totalPengeluaranLabaRugi = 0;
var totalPendapatanLabaRugi = 0;

var tz = 'Asia/Jakarta';
for (var d = 6; d >= 0; d--) {
var dateObj = new Date();
dateObj.setDate(dateObj.getDate() - d);
var label = Utilities.formatDate(dateObj, tz, 'dd/MM');
var fullLabel = Utilities.formatDate(dateObj, tz, 'dd/MM/yyyy');
trendMap[fullLabel] = { label: label, Pemasukan: 0, Pengeluaran: 0 };
}

for (var i = 1; i < data.length; i++) {
var dateStr = data[i][1];
var tipe = data[i][3];
var kategori = data[i][4];
var nominal = parseFormattedNumber(data[i][5]);

if (trendMap[dateStr]) {
trendMap[dateStr][tipe] += nominal;
}

if (tipe === 'Pengeluaran') {
totalPengeluaranLabaRugi += nominal;
if (!expenseCategoryMap[kategori]) {
expenseCategoryMap[kategori] = 0;
}
expenseCategoryMap[kategori] += nominal;
} else if (tipe === 'Pemasukan') {
totalPendapatanLabaRugi += nominal;
if (!incomeCategoryMap[kategori]) {
incomeCategoryMap[kategori] = 0;
}
incomeCategoryMap[kategori] += nominal;
}
}

// Breakdown Pengeluaran
var breakdownPengeluaran = [];
for (var cat in expenseCategoryMap) {
var amt = expenseCategoryMap[cat];
var pct = totalPengeluaranLabaRugi > 0 ? (amt / totalPengeluaranLabaRugi) * 100 : 0;
breakdownPengeluaran.push({
kategori: cat,
nominal: amt,
persentase: pct.toFixed(2)
});
}
breakdownPengeluaran.sort(function(a, b) { return b.nominal - a.nominal; });

// Breakdown Pendapatan
var breakdownPendapatan = [];
for (var incCat in incomeCategoryMap) {
var incAmt = incomeCategoryMap[incCat];
var incPct = totalPendapatanLabaRugi > 0 ? (incAmt / totalPendapatanLabaRugi) * 100 : 0;
breakdownPendapatan.push({
kategori: incCat,
nominal: incAmt,
persentase: incPct.toFixed(2)
});
}
breakdownPendapatan.sort(function(a, b) { return b.nominal - a.nominal; });

var labelsTrend = [];
var dataIn = [];
var dataOut = [];
for (var k in trendMap) {
labelsTrend.push(trendMap[k].label);
dataIn.push(trendMap[k].Pemasukan);
dataOut.push(trendMap[k].Pengeluaran);
}

return {
totalPengeluaran: totalPengeluaranLabaRugi,
totalPendapatan: totalPendapatanLabaRugi,
breakdownKategori: breakdownPengeluaran,
breakdownPendapatan: breakdownPendapatan,
trend: {
labels: labelsTrend,
pemasukan: dataIn,
pengeluaran: dataOut
}
};
} catch(e) {
return { error: e.toString() };
}
}

// 5. SAVE TRANSACTION API
function saveTransaction(tipe, kategori, nominal, catatan) {
try {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('transaksi');

var now = new Date();
var dateStr = Utilities.formatDate(now, 'Asia/Jakarta', 'dd/MM/yyyy');
var timeStr = Utilities.formatDate(now, 'Asia/Jakarta', 'HH:mm:ss');

var txId = generateTransactionId(dateStr);
var cleanNominal = parseFormattedNumber(nominal);

sheet.appendRow([txId, dateStr, timeStr, tipe, kategori, cleanNominal, catatan]);
SpreadsheetApp.flush();

return { success: true, id: txId, tanggal: dateStr, jam: timeStr };
} catch(e) {
return { success: false, error: e.toString() };
}
}

// 6. CATEGORY MANAGEMENT ACTIONS
function getCategories() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('kategori');
var data = sheet.getDataRange().getDisplayValues();
var categories = [];
for (var i = 1; i < data.length; i++) {
if (data[i][0]) {
categories.push({
id: data[i][0],
tipe: data[i][1],
nama: data[i][2]
});
}
}
return categories;
}

function saveCategory(tipe, nama) {
try {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('kategori');
var data = sheet.getDataRange().getDisplayValues();

var maxId = 0;
for (var i = 1; i < data.length; i++) {
var idStr = data[i][0];
if (idStr && idStr.indexOf('KAT-') === 0) {
var num = parseInt(idStr.substring(4), 10);
if (!isNaN(num) && num > maxId) maxId = num;
}
}
var nextId = "KAT-" + ("000" + (maxId + 1)).slice(-4);

sheet.appendRow([nextId, tipe, nama]);
SpreadsheetApp.flush();
return { success: true, id: nextId, tipe: tipe, nama: nama };
} catch(e) {
return { success: false, error: e.toString() };
}
}

function deleteCategory(id) {
try {
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('kategori');
var data = sheet.getDataRange().getDisplayValues();

for (var i = 1; i < data.length; i++) {
if (data[i][0] === id) {
sheet.deleteRow(i + 1);
SpreadsheetApp.flush();
return { success: true };
}
}
return { success: false, error: 'Kategori tidak ditemukan.' };
} catch(e) {
return { success: false, error: e.toString() };
}
}