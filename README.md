//kredensial

const spreadsheetId   = '****'
const sheetName       = 'data1' // Ini akan digunakan sebagai sheet data utama Anda
const logSheetName    = 'Log'

const botHandle       = '***'
const botToken        = '****'
const appsScriptUrl   = '***'
const telegramApiUrl  = `https://api.telegram.org/bot${botToken}` // Menggunakan botToken yang sudah didefinisikan

// ---

function log(logMessage = '') {
  // akses sheet
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId)
  const sheet       = spreadsheet.getSheetByName(logSheetName)
  const lastRow     = sheet.getLastRow()
  const row         = lastRow + 1

  // inisiasi nilai
  const today = new Date

  // insert row kosong
  sheet.insertRowAfter(lastRow)

  // insert data
  sheet.getRange(`A${row}`).setValue(today)
  sheet.getRange(`B${row}`).setValue(logMessage)
}

function sendTelegramMessage(chatId, replyToMessageId, textMessage) {
  // url kirim pesan
  const url = `${telegramApiUrl}/sendMessage`;
  
  // payload
  const data = {
    parse_mode            : 'HTML',
    chat_id               : chatId,
    reply_to_message_id   : replyToMessageId,
    text                  : textMessage,
    disable_web_page_preview: true,
  }
  
  const options = {
    method     : 'post',
    contentType: 'application/json',
    payload    : JSON.stringify(data)
  }

  try {
    const response = UrlFetchApp.fetch(url, options).getContentText()
    return response;
  } catch (e) {
    log(`Error sending Telegram message to ${chatId}: ${e.message}`);
    return null; // Mengembalikan null jika ada error
  }
}


function formatDate(date) {
  const monthIndoList = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Ags', 'Sep', 'Okt', 'Nov', 'Des']

  const dateIndo  = date.getDate()
  const monthIndo = monthIndoList[date.getMonth()]
  const yearIndo  = date.getFullYear()

  const result = `${dateIndo} ${monthIndo} ${yearIndo}`

  return result
}


function parseMessage(message = '') {
  // pisahkan berdasarkan karakter enter
  const splitted = message.split('\n')

  // inisiasi variabel
  let nama       = ''
  let kodeBarang = ''
  let alamat     = ''
  let resi       = ''

  // parsing pesan untuk mencari nilai variabel
  splitted.forEach(el => {
    nama = el.includes('Nama:') ? el.split(':')[1].trim().replaceAll('\n', ' ') : nama;
    kodeBarang = el.includes('Kode Barang:') ? el.split(':')[1].trim().replaceAll('\n', ' ') : kodeBarang;
    alamat = el.includes('Alamat:') ? el.split(':')[1].trim().replaceAll('\n', ' ') : alamat;
    resi = el.includes('Resi:') ? el.split(':')[1].trim().replaceAll('\n', ' ') : resi;
  })

  // kumpulkan hasil
  const result = {
    nama       : nama,
    kodeBarang : kodeBarang,
    alamat     : alamat,
    resi       : resi,
  }

  // jika data kosong
  const isEmpty = (nama === '' && kodeBarang === '' && alamat === '' && resi === '')

  return isEmpty ? false : result
}


function inputDataOrder(data) {
  try {
    // akses sheet
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId)
    // *** PERBAIKAN: Menggunakan sheetName yang sudah didefinisikan ***
    const sheet = spreadsheet.getSheetByName(sheetName) 
    const lastRow = sheet.getLastRow()
    const row = lastRow + 1

    // inisiasi nilai
    // *** PERBAIKAN: Menggunakan row sebagai nomor urut, atau lastRow + 1 jika ingin dimulai dari 1 ***
    // Jika baris pertama data adalah A2, maka lastRow akan 1 (header), jadi nomor urut bisa lastRow.
    // Jika sheet kosong, lastRow = 0, jadi number = 0. Jika ingin dimulai dari 1, gunakan row.
    const number  = lastRow > 0 ? lastRow : 0; // Nomor urut berdasarkan baris terakhir
    const idOrder = `ORD-${number + 1}` // *** PERBAIKAN: +1 agar ID dimulai dari ORD-1 jika sheet kosong atau ORD-X+1 dari data terakhir ***
    const today   = new Date

    // insert row kosong
    // *** PERBAIKAN: Jika sheet kosong, insertRowAfter(0) akan error. Gunakan appendRow jika ingin aman ***
    // Namun, jika Anda selalu ingin menyisipkan di antara data atau mengelola header, insertRowAfter OK.
    // Asumsi baris 1 adalah header, jadi data dimulai dari baris 2.
    // getLastRow() akan mengembalikan 1 jika hanya ada header.
    // Jika Anda ingin data baru selalu di baris berikutnya tanpa header, baris ini sudah benar.
    // Jika Anda ingin data dimulai dari baris 2 dan getLastRow() menghasilkan 0 jika kosong, Anda harus tangani kasus ini.
    // Untuk lebih aman, bisa gunakan:
    if (lastRow === 0) { // Jika sheet kosong (belum ada data, mungkin hanya header)
        // Tidak perlu insertRowAfter, langsung set value
    } else {
        sheet.insertRowAfter(lastRow);
    }

    // insert data
    sheet.getRange(`A${row}`).setValue(number + 1) // *** PERBAIKAN: Menyimpan nomor urut yang benar
    sheet.getRange(`B${row}`).setValue(idOrder)
    sheet.getRange(`C${row}`).setValue(today)
    sheet.getRange(`D${row}`).setValue(data['nama'])
    sheet.getRange(`E${row}`).setValue(data['kodeBarang'])
    sheet.getRange(`F${row}`).setValue(data['alamat'])
    sheet.getRange(`G${row}`).setValue(data['resi'])
    sheet.getRange(`H${row}`).setValue('Sedang dikemas')
    sheet.getRange(`I${row}`).setValue(data['chatId'])

    // jika berhasil, return idOrder
    return idOrder

  } catch(err) {
    log(`Error inputDataOrder: ${err.message}. Data: ${JSON.stringify(data)}`); // Log error
    return false
  }
}


function cekResi(resi = null) {
  // cegah resi kosong
  if (!resi) {
    return 'Format pencarian resi tidak valid.'
  }

  // akses sheet
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId)
  // *** PERBAIKAN: Menggunakan sheetName yang sudah didefinisikan ***
  const sheet       = spreadsheet.getSheetByName(sheetName)
  const lastRow     = sheet.getLastRow()

  // ambil data
  // Pastikan Anda memiliki header di baris 1, jika tidak, mulai dari A1.
  // Jika sheet kosong (lastRow=0), range ini akan menyebabkan error. Tangani kasus ini.
  if (lastRow < 2) { // Asumsi baris 1 adalah header, jadi data dimulai dari baris 2
      return `Resi ${resi} tidak ditemukan (sheet kosong atau hanya header).`;
  }
  const range       = `A2:I${lastRow}`
  const dataList    = sheet.getRange(range).getValues()

  // filter data
  const dataListFiltered = dataList.filter(el => el[6].toString().toLowerCase() === resi.toString().toLowerCase())

  // cek jika resi ditemukan   
  const isResiFound = dataListFiltered.length > 0

  // variabel balasan
  let messageReply = ''

  // jika ditemukan
  if (isResiFound) {
    // jika ada no resi yang sama, yang diambil yang paling atas
    const data = dataListFiltered[0]

    messageReply = `Info Resi <b>${resi}</b>

ID Order: ${data[1]}
Tanggal Order: ${formatDate(data[2])}
Nama: ${data[3]}
Kode Barang: ${data[4]}
Alamat: ${data[5]}
Status Pengiriman: <b>${data[7]}</b>`
  
  // jika tidak
  } else {
    messageReply = `Resi ${resi} tidak ditemukan.`
  }

  return messageReply
}


function handleUpdateDeliveryStatus(e) {
  // ambil info sheet dan row yang baru diedit
  const row         = e.range.getRow()
  const column      = e.range.getA1Notation().replace(/[^a-zA-Z]/g, '')
  const sheetNameTrigger = e.range.getSheet().getSheetName() // Menggunakan nama variabel berbeda agar tidak bentrok dengan global sheetName

  // jika perubahan bukan pada sheet data order kolom H
  // *** PERBAIKAN: Menggunakan sheetName yang sudah didefinisikan ***
  if (sheetNameTrigger !== sheetName || column !== 'H') {
    return false
  }

  // akses sheet
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId)
  // *** PERBAIKAN: Menggunakan sheetName yang sudah didefinisikan ***
  const sheet       = spreadsheet.getSheetByName(sheetName)
  const today       = new Date

  // ambil data
  const range = `A${row}:I${row}`
  const data  = sheet.getRange(range).getValues()

  // isi konstanta
  const idOrder          = data[0][1]
  const tanggalOrder     = data[0][2]
  const nama             = data[0][3]
  const kodeBarang       = data[0][4]
  const alamat           = data[0][5]
  const resi             = data[0][6]
  const statusPengiriman = data[0][7]
  const chatId           = data[0][8].toString() // Pastikan chatId adalah string

  const textMessage = `Update Info Resi <b>${resi}</b>

ID Order: ${idOrder}
Tanggal Order: ${formatDate(tanggalOrder)}
Nama: ${nama}
Kode Barang: ${kodeBarang}
Alamat: ${alamat}
Status Pengiriman: <b>${statusPengiriman}</b>

<i>Data per-${formatDate(today)}</i>`

  // kirim pesan
  sendTelegramMessage(chatId, null, textMessage)
}


function doPost(e) {
  try {
    // urai pesan masuk
    const contents          = JSON.parse(e.postData.contents)
    const chatId            = contents.message.chat.id
    const receivedTextMessage = contents.message.text ? contents.message.text.replace(botHandle, '').trim() : ''; // Tangani jika text null
    const messageId           = contents.message.message_id

    let messageReply = ''

    // 1. jika pesan /start
    if (receivedTextMessage.toLowerCase() === '/start') {
      // tulis pesan balasan
      messageReply = `Halo! Status bot dalam keadaan aktif.`

    // 2. jika pesan diawali dengan /input
    } else if (receivedTextMessage.split('\n')[0].toLowerCase() === '/input') {
      const parsedMessage = parseMessage(receivedTextMessage)

      // 2a.jika ada data
      if (parsedMessage) {
        const data = {
          nama       : parsedMessage['nama'],
          kodeBarang : parsedMessage['kodeBarang'],
          alamat     : parsedMessage['alamat'],
          resi       : parsedMessage['resi'],
          chatId     : chatId // Tambahkan chatId dari pesan Telegram
        }

        // insert data ke sheet
        const idOrder = inputDataOrder(data)

        // tulis pesan balasan
        messageReply = idOrder ? `Data berhasil disimpan dengan ID Order <b>${idOrder}</b>` : 'Data gagal disimpan'

      // 2b. jika tidak ada data
      } else {
        messageReply = 'Data kosong dan tidak dapat disimpan'
      }

    // 3. cek resi 
    } else if (receivedTextMessage.split(' ')[0].toLowerCase() === '/resi') {
      // ambil resi
      const resi = receivedTextMessage.split(' ')[1]

      // ambil info
      messageReply = cekResi(resi)

    // 4. format
    } else if (receivedTextMessage.toLowerCase() === '/format') {
      messageReply = `Untuk <b>input data order</b> gunakan format:

<pre>/input
Nama: 
Kode Barang: 
Alamat: 
Resi: </pre>

Untuk <b>cek resi</b> gunakan format:

<pre>/resi [nomor resi]</pre>
(Tanpa tanda kurung siku)`

    // 5. format salah
    } else {
      messageReply = `Pesan yang Anda kirim tidak sesuai format.

Kirim perintah /format untuk melihat daftar format pesan yang tersedia.`
    }

    // kirim pesan balasan
    sendTelegramMessage(chatId, messageId, messageReply)

  } catch(err) {
    log(`Error in doPost: ${err.message}. Request data: ${e.postData.contents}`); // Log error lebih detail
    sendTelegramMessage(chatId, null, `Terjadi error: ${err.message}. Silakan coba lagi atau hubungi admin.`); // Beri feedback ke user Telegram
  }
}


function setWebhook() {
  // akses api
  const url        = `${telegramApiUrl}/setwebhook?url=${appsScriptUrl}`
  const response = UrlFetchApp.fetch(url).getContentText()
  
  Logger.log(response)
}
