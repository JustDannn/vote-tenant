/**
 * Backend vote The Market Day 2026 (TUKS).
 *
 * Cara pakai:
 * 1. Buat Google Sheet baru -> Extensions -> Apps Script, tempel file ini.
 * 2. Jalankan fungsi `setup` sekali (izinkan akses) untuk membuat sheet Votes & Rekap.
 * 3. Deploy -> New deployment -> Web app
 *    Execute as: Me, Who has access: Anyone -> salin URL /exec-nya.
 *
 * Request dari VotePage.vue: POST body JSON { googleId, email, name, tenantId }
 * Response: { success: boolean, message: string }
 */

const VOTES_SHEET = 'Votes'
const REKAP_SHEET = 'Rekap'

// Disalin dari dataset-tenant.csv. Ubah juga di sini kalau dataset berubah.
const TENANTS = [
  ["T01", "Sleepy Ruben", "TMD-01"],
  ["T02", "Bite N' Brew", "TMD-02"],
  ["T03", "DaPreneur", "TMD-03"],
  ["T04", "Dolo & Cake", "TMD-04"],
  ["T05", "Noury", "TMD-05"],
  ["T06", "Kalamoy", "TMD-06"],
  ["T07", "KeyPi", "TMD-07"],
  ["T08", "Toastclub", "TMD-08"],
  ["T09", "Strappy Shop", "TMD-09"],
  ["T10", "Charmora", "TMD-10"],
  ["T11", "Kameraku", "TMD-11"],
  ["T12", "Vyre", "TMD-12"],
  ["T13", "Charmera", "TMD-13"],
  ["T14", "Wangy Rek", "TMD-14"],
  ["T15", "FROSTee", "TMD-15"],
  ["T16", "Cleanver", "TMD-16"],
  ["T17", "DecooLabs", "TMD-17"],
  ["T18", "kiracha.id", "TMD-18"],
  ["T19", "Yummiou", "TMD-19"],
  ["T20", "Socks.IN", "TMD-20"],
  ["T21", "SevCare", "TMD-21"],
  ["T22", "KitschClub", "TMD-22"],
  ["T23", "MyToken", "TMD-23"],
  ["T24", "jajanOsek", "TMD-24"],
  ["T25", "MARIMARI", "TMD-25"],
  ["T26", "CIRO", "TMD-26"],
  ["T27", "Kedai Atji Akang", "TMD-27"],
  ["T28", "mangolicious", "TMD-28"],
  ["T29", "DONACHA", "TMD-29"],
  ["T30", "Fruity Rush", "TMD-30"],
  ["T31", "Milkie", "TMD-31"],
  ["T32", "Rebelcrunchh", "TMD-32"],
  ["T33", "MBG (MIE BAKAR GANK)", "TMD-33"],
  ["T34", "KWU ROAD TO 1 MILLION", "TMD-34"],
  ["T35", "COMTERFOOD INDONESIA", "TMD-35"],
  ["T36", "LAB LAB DORRRR", "TMD-36"],
  ["T37", "Doughsy", "TMD-37"],
  ["T38", "GABELO ICECREAM", "TMD-38"],
  ["T39", "FruRolls", "TMD-39"],
  ["T40", "TetesDrink", "TMD-40"],
  ["T41", "PUD & CO", "TMD-41"],
  ["T42", "CHICKEPE", "TMD-42"],
  ["T43", "K3 GEMBIRA", "TMD-43"],
  ["T44", "Slack's", "TMD-44"],
  ["T45", "SEPRINT", "TMD-45"],
  ["T46", "Joss Jiss", "TMD-46"],
  ["T47", "QBITEZ", "TMD-47"],
  ["T48", "TEHANAN", "TMD-48"],
  ["T49", "MOCHICHUU", "TMD-49"],
  ["T50", "Brow&Grit", "TMD-50"],
  ["T51", "Nyeruput Nyemil", "TMD-51"],
  ["T52", "9MOOD", "TMD-52"],
]

function doPost(e) {
  const lock = LockService.getScriptLock()
  try {
    lock.waitLock(10000)

    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}')
    const googleId = String(body.googleId || '').trim()
    const email = String(body.email || '').trim().toLowerCase()
    const name = String(body.name || '').trim()
    const tenantId = String(body.tenantId || '').trim().toUpperCase()

    if (!googleId || !email || !tenantId) {
      return json({ success: false, message: 'Data vote tidak lengkap. Silakan login ulang.' })
    }

    const tenant = TENANTS.find((t) => t[0] === tenantId)
    if (!tenant) {
      return json({ success: false, message: `Tenant ${tenantId} tidak terdaftar.` })
    }

    const sheet = getVotesSheet()
    const lastRow = sheet.getLastRow()
    if (lastRow > 1) {
      // Kolom B (googleId) s.d. E (tenantId)
      const rows = sheet.getRange(2, 2, lastRow - 1, 4).getValues()
      const alreadyVoted = rows.some((row) => String(row[0]) === googleId && row[3] === tenantId)
      if (alreadyVoted) {
        return json({ success: false, message: `Kamu sudah vote ${tenant[1]}. Satu akun hanya bisa vote sekali per tenant.` })
      }
    }

    sheet.appendRow([new Date(), googleId, email, name, tenantId, tenant[1]])
    return json({ success: true, message: `Vote untuk ${tenant[1]} berhasil tercatat. Terima kasih!` })
  } catch (err) {
    return json({ success: false, message: 'Gagal memproses vote: ' + err.message })
  } finally {
    lock.releaseLock()
  }
}

function doGet() {
  return json({ success: true, message: 'Vote backend aktif.' })
}

function setup() {
  getVotesSheet()

  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const rekap = ss.getSheetByName(REKAP_SHEET) || ss.insertSheet(REKAP_SHEET)
  rekap.clear()
  rekap.getRange(1, 1, 1, 4).setValues([['tenant_id', 'nama_tenant', 'booth', 'jumlah_vote']]).setFontWeight('bold')
  const rows = TENANTS.map((t, i) => [t[0], t[1], t[2], `=COUNTIF(${VOTES_SHEET}!E:E, A${i + 2})`])
  rekap.getRange(2, 1, rows.length, 4).setValues(rows)
  rekap.setFrozenRows(1)
  rekap.autoResizeColumns(1, 4)
}

function getVotesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  let sheet = ss.getSheetByName(VOTES_SHEET)
  if (!sheet) {
    sheet = ss.insertSheet(VOTES_SHEET)
    sheet.getRange(1, 1, 1, 6)
      .setValues([['timestamp', 'google_id', 'email', 'name', 'tenant_id', 'nama_tenant']])
      .setFontWeight('bold')
    sheet.setFrozenRows(1)
  }
  return sheet
}

function json(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON)
}
