function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://cdn-icons-png.flaticon.com/512/8853/8853008.png')
    .setTitle('Portfolio')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getLoginData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Login');
  return sheet.getRange('A2:C' + sheet.getLastRow()).getValues().map(row => ({
    username: row[0],
    password: row[1],
    namaLengkap: row[2] 
  }));
}

function onEdit(e) {
  const ss = e.source;
  const loginSheet = ss.getSheetByName('Login');
  const identitasSheet = ss.getSheetByName('Identitas');
  if (!loginSheet || !identitasSheet) return;
  const editedSheet = e.range.getSheet();
  const editedRow = e.range.getRow();
  const editedColumn = e.range.getColumn();
  if (editedSheet.getName() === 'Login' && editedColumn === 3 && editedRow >= 2) {
    const loginData = loginSheet.getRange('C2:C' + loginSheet.getLastRow()).getValues();
    if (identitasSheet.getLastRow() >= 2) {
      identitasSheet.getRange(2, 2, identitasSheet.getLastRow() - 1, 1).clearContent();
    }
    if (loginData.length > 0) {
      identitasSheet.getRange(2, 2, loginData.length, 1).setValues(loginData);
    }
  }
}


function getDropdownData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Login');
  const data = sheet.getRange('C2:C' + sheet.getLastRow()).getValues();
  return data.flat().map(item => item.toString().trim()).filter(item => item);
}

function getGuruData(nama) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Identitas');
  const data = sheet.getRange('A2:C' + sheet.getLastRow()).getValues();
  const cleanName = nama ? nama.trim().toLowerCase() : "";
  const guru = data.find(row => row[1].toString().trim().toLowerCase() === cleanName);
  if (guru) {
    let fotoUrl = guru[0];
    const match = guru[0].match(/[-\w]{25,}/);
    if (match) {
      fotoUrl = `https://lh3.googleusercontent.com/d/${match[0]}`;
    }
    return {
      nama: guru[1],
      nip: guru[2],
      foto: fotoUrl
    };
  }
  return null;
}

function getDashboardStats(namaGuru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pengalaman = getRowCount(ss, 'Pengalaman', 'A2:C', namaGuru); 
  const mapel = getRowCount(ss, 'Pelajaran', 'A2:C', namaGuru); 
  const siswa = getRowCount(ss, 'Siswa', 'A2:F', namaGuru); 
  const administrasi = getRowCount(ss, 'Administrasi', 'A2:G', namaGuru); 
  return {
    pengalaman: pengalaman,
    mapel: mapel,
    siswa: siswa,
    administrasi: administrasi
  };
}

function getRowCount(ss, sheetName, range, namaGuru) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0; 
  const data = sheet.getRange(range).getValues();
  return data.filter(row => row.some(cell => cell !== "") && row.includes(namaGuru)).length;
}

function getIdentitas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Identitas');
  const data = sheet.getRange('A2:C').getValues().filter(row => row[0]);
  return data.map(row => {
    const driveId = row[0].match(/[-\w]{25,}/); 
    return {
      foto: driveId ? `https://lh3.googleusercontent.com/d/${driveId[0]}` : '',
      nama: row[1],
      nip: row[2]
    };
  });
}

function getPengalamanGuru(namaGuru) {
  if (!namaGuru) {
    throw new Error("Parameter namaGuru kosong atau undefined");
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pengalaman');
  const data = sheet.getRange('B2:C' + sheet.getLastRow()).getValues();
  const pengalaman = data
    .filter(row => row[1] && row[1].trim().toLowerCase() === namaGuru.trim().toLowerCase())
    .map((row, index) => ({
      idPeng: index + 1,
      deskripsi: row[0]
    }));
  return pengalaman;
}

function addPengalaman(pengalaman, namaGuru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pengalaman');
  const ids = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat();
  const lastId = ids.length > 0 ? Math.max(...ids.filter(id => !isNaN(id))) : 0;
  sheet.appendRow([lastId + 1, pengalaman, namaGuru]);
  return 'Pengalaman berhasil ditambahkan!';
}

function getPengalamanByNama(namaLengkap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pengalaman');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  return data
    .filter(row => row[2] === namaLengkap)
    .map(row => [row[0], row[1]]);
}

function updatePengalaman(id, pengalamanBaru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pengalaman');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  const rowIndex = data.findIndex(row => row[0] == id);
  if (rowIndex !== -1) {
    sheet.getRange(rowIndex + 2, 2).setValue(pengalamanBaru);
    return "Pengalaman berhasil diperbarui!";
  } else {
    return "ID tidak ditemukan!";
  }
}

function deletePengalaman(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pengalaman');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  const rowIndex = data.findIndex(row => row[0] == id);
  if (rowIndex !== -1) {
    sheet.deleteRow(rowIndex + 2); 
    return "Pengalaman berhasil dihapus!";
  } else {
    return "ID tidak ditemukan!";
  }
}

function getPelajaran(namaGuru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Pelajaran');
  const data = sheet.getRange('A2:C').getValues();
  return data
    .filter(row => row[2] === namaGuru) 
    .map(row => ({
      idPel: row[0],
      keterangan: row[1]
    }));
}

function getMapelByGuru(namaGuru) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pelajaran');
  if (!sheet) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  return data
    .filter(row => row[2] === namaGuru)
    .map(row => ({
      id: row[0],  
      mapel: row[1] 
    }));
}

function addMapel(mapel, namaGuru) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pelajaran');
  const ids = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat();
  const lastId = ids.length > 0 ? Math.max(...ids.filter(id => !isNaN(id))) : 0;
  sheet.appendRow([lastId + 1, mapel, namaGuru]);
}

function updateMapel(id, mapel) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pelajaran');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  const rowIndex = data.findIndex(row => row[0] == id);
  if (rowIndex !== -1) {
    sheet.getRange(rowIndex + 2, 2).setValue(mapel);
  } else {
    throw new Error('ID tidak ditemukan');
  }
}

function deleteMapel(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pelajaran');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  const rowIndex = data.findIndex(row => row[0] == id);
  if (rowIndex !== -1) {
    sheet.deleteRow(rowIndex + 2);
  } else {
    throw new Error('ID tidak ditemukan');
  }
}

function getSiswa(namaGuru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Siswa');
  const data = sheet.getRange('A2:F').getValues().filter(row => row[0]);
  const filteredData = data.filter(row => row[5] === namaGuru);
  return filteredData.map(row => ({
    nama: row[1],
    nis: row[2],
    gender: row[3],
    kelas: row[4]
  }));
}

function getSiswaByGuru(namaGuru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Siswa');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  return data
    .map((row, index) => ({
      id: row[0],  
      nama: row[1], 
      nis: row[2],  
      gender: row[3], 
      kelas: row[4], 
      guru: row[5], 
      rowIndex: index + 2 
    }))
    .filter(item => item.guru === namaGuru);
}

function addSiswaData(nama, nis, gender, kelas, namaGuru) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Siswa');
    const ids = sheet.getRange('A2:A' + sheet.getLastRow()).getValues().flat();
    const lastId = ids.length > 0 ? Math.max(...ids.filter(id => !isNaN(id))) : 0;
    sheet.appendRow([lastId + 1, nama, nis, gender, kelas, namaGuru]);
}

function updateSiswaById(id, nama, nis, gender, kelas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Siswa');
  const data = sheet.getRange('A2:A' + sheet.getLastRow()).getValues();
  const rowIndex = data.findIndex(row => row[0] == id) + 2;
  if (rowIndex >= 2) {
    const namaGuru = sheet.getRange(rowIndex, 6).getValue();
    sheet.getRange(rowIndex, 2, 1, 5).setValues([[nama, nis, gender, kelas, namaGuru]]);
    return 'Sukses';
  } else {
    throw new Error('ID tidak ditemukan.');
  }
}

function deleteSiswaById(id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Siswa');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues(); 
  for (let i = 0; i < data.length; i++) {
    if (data[i][0].toString() === id.toString()) {
      sheet.deleteRow(i + 2);
      return 'Berhasil';
    }
  }
  return 'ID tidak ditemukan';
}

function getResources(namaGuru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Administrasi');
  const data = sheet.getRange('A2:G').getValues().filter(row => row[0]);
  const filteredData = data.filter(row => row[6] === namaGuru && row[5] === 'Aktif');
  return filteredData.map((row, index) => ({
    no: index + 1,
    nama: row[1],
    mapel: row[2],
    semester: row[3],
    url: row[4]
  }));
}

function getAdministrasiByGuru(namaGuru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Administrasi');
  if(!sheet || sheet.getLastRow() < 2) return [];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  return data.filter(row => row[6] === namaGuru);
}

function uploadAdministrasiData(fileData) {
    try {
        const folderName = 'ASSET';
        const folderIterator = DriveApp.getFoldersByName(folderName);
        const folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.createFolder(folderName);
        if (!fileData || !fileData.bytes || !fileData.mimeType || !fileData.fileName) {
            throw new Error("Data file tidak valid atau tidak lengkap.");
        }
        const blob = Utilities.newBlob(fileData.bytes, fileData.mimeType, fileData.fileName);
        const upload = folder.createFile(blob);
        const url = upload.getUrl();
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Administrasi');
        if (!sheet) throw new Error("Sheet 'Administrasi' tidak ditemukan.");
        const lastRow = sheet.getLastRow();
        const existingIds = lastRow > 1
            ? sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat()
            : [];
        function generateUniqueId() {
            let id;
            do {
                id = Math.floor(10000 + Math.random() * 90000); // ID unik 5 digit
            } while (existingIds.includes(id));
            return id;
        }
        const uniqueId = generateUniqueId();
        sheet.appendRow([
            uniqueId,
            fileData.nama || "Tidak Ada Nama",
            fileData.mapel || "Tidak Ada Mapel",
            fileData.semester || "Tidak Ada Semester",
            url,
            fileData.status || "Pending",
            fileData.guru || "Tidak Ada Guru"
        ]);
        return `File berhasil diunggah: ${url}`;

    } catch (error) {
        Logger.log("Error di uploadAdministrasiData: " + error.message);
        throw new Error("Gagal mengunggah data administrasi: " + error.message);
    }
}


function updateAdministrasiData(id, nama, mapel, semester, fileData, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Administrasi');
  const data = sheet.getDataRange().getValues();
  const idColumn = 0; 
  const rowIndex = data.findIndex(row => row[idColumn] == id);
  if (rowIndex === -1) return; 
  let fileUrl = data[rowIndex][4]; 
  if (fileData) {
    const folderName = 'ASSET';
    const folderIterator = DriveApp.getFoldersByName(folderName);
    const folder = folderIterator.hasNext() ? folderIterator.next() : DriveApp.createFolder(folderName);
    if (fileUrl) {
      const fileId = fileUrl.split('/d/')[1].split('/')[0];
      try {
        DriveApp.getFileById(fileId).setTrashed(true);
      } catch (e) {
        Logger.log('File tidak ditemukan atau sudah dihapus.');
      }
    }
    const blob = Utilities.newBlob(fileData.bytes, fileData.mimeType, fileData.fileName);
    const newFile = folder.createFile(blob);
    fileUrl = newFile.getUrl();
  }
  sheet.getRange(rowIndex + 1, 2, 1, 5).setValues([[nama, mapel, semester, fileUrl, status]]);
}


function deleteAdministrasiById(id) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Administrasi');
  const data = sheet.getDataRange().getValues();
  let rowToDelete = -1;
  let fileUrl = '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      rowToDelete = i + 1;
      fileUrl = data[i][4];
      break;
    }
  }
  if (rowToDelete !== -1) {
    if (fileUrl) {
      const fileId = fileUrl.match(/[-\w]{25,}/);
      if (fileId) {
        try {
          const file = DriveApp.getFileById(fileId[0]);
          file.setTrashed(true); // Memindahkan ke sampah
        } catch (e) {
          Logger.log("File tidak ditemukan atau sudah dihapus.");
        }
      }
    }
    sheet.deleteRow(rowToDelete);
    return true;
  }
  return false;
}

function saveContactData(nama, subjek, pesan) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Kontak');
  if (!sheet) {
    sheet = ss.insertSheet('Kontak');
    sheet.appendRow(['Nama', 'Subjek', 'Pesan', 'Tanggal']);
  }
  const tanggal = "'" + new Date().toLocaleString('id-ID');
  sheet.appendRow([nama, subjek, pesan, tanggal]);
}

function getPortfolioData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Identitas');
  const data = sheet.getRange('A2:C' + sheet.getLastRow()).getValues();
  return data;
}

function updatePortfolioWithImage(namaLama, namaBaru, nipBaru, imageBase64) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Identitas');
  const data = sheet.getRange('A2:C' + sheet.getLastRow()).getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][1] === namaLama) {
      let imageUrl = data[i][0]; 
      if (imageBase64) {
        const folder = getOrCreateFolder('ASSET');
        const blob = Utilities.newBlob(Utilities.base64Decode(imageBase64), "image/png", namaBaru + ".png");
        const file = folder.createFile(blob);
        imageUrl = file.getUrl(); 
      }
      sheet.getRange(i + 2, 1).setValue(imageUrl); 
      sheet.getRange(i + 2, 2).setValue(namaBaru); 
      sheet.getRange(i + 2, 3).setValue(nipBaru);
      return "Data berhasil diperbarui!";
    }
  }
  return "Data tidak ditemukan!";
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
}

function getKontakData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Kontak');
  const data = sheet.getRange('A2:D' + sheet.getLastRow()).getValues();
  return data.map((row, index) => ({
    index: index + 2, 
    nama: row[0],
    subjek: row[1],
    pesan: row[2],
    timestamp: row[3]
  }));
}

function deleteKontakRow(rowIndex) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kontak');
  sheet.deleteRow(rowIndex);
}


