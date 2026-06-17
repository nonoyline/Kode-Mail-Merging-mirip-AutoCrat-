/**
 * Google Apps Script - Autocrat Replacer Wizard
 * 
 * Script ini menyediakan panel visual (Sidebar) berbentuk Wizard langkah-demi-langkah
 * untuk melakukan merge dokumen tingkat lanjut (mirip dengan Autocrat).
 * 
 * CARA PENGGUNAAN:
 * 1. Buka Google Sheets Anda.
 * 2. Klik Ekstensi (Extensions) > Apps Script.
 * 3. Hapus kode default, lalu salin (paste) seluruh kode script ini ke sana.
 * 4. Klik Simpan (ikon disket).
 * 5. Muat ulang (refresh) Google Sheets Anda.
 * 6. Menu baru bernama "Merge Dokumen" akan muncul di bilah menu atas Google Sheets.
 * 7. Klik "Merge Dokumen" > "Buka Panel Konfigurasi" untuk mulai menggunakan.
 */

// ==========================================
// 1. PEMBUATAN MENU SPREADSHEET
// ==========================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Merge Dokumen')
    .addItem('Buka Wizard Konfigurasi', 'showSidebar')
    .addToUi();
}

// ==========================================
// 2. MENAMPILKAN SIDEBAR UI
// ==========================================
function showSidebar() {
  const htmlContent = doGetSidebarHtml();
  const htmlOutput = HtmlService.createHtmlOutput(htmlContent)
    .setTitle('Autocrat Replacer Wizard')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

// ==========================================
// 3. BACKEND API UNTUK SPREADSHEET & DRIVE
// ==========================================

/**
 * Mendapatkan daftar semua sheet di Spreadsheet aktif
 */
function getSpreadsheetSheets() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s => s.getName());
}

/**
 * Mendapatkan daftar kolom header (baris 1) dari sheet tertentu
 */
function getSheetHeaders(sheetName) {
  if (!sheetName) return [];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getDisplayValues();
  if (data.length === 0) return [];
  return data[0].map(h => h.toString().trim()).filter(h => h !== '');
}

/**
 * Mencari dokumen Google Doc atau Folder di Google Drive secara real-time
 */
function searchDriveFiles(query, isFolder) {
  const results = [];
  let mimeTypeCondition = isFolder 
    ? "mimeType = 'application/vnd.google-apps.folder'" 
    : "mimeType = 'application/vnd.google-apps.document'";
  let searchString = mimeTypeCondition + " and trashed = false";
  
  if (query) {
    const cleanQuery = query.replace(/'/g, "\\'");
    searchString += " and title contains '" + cleanQuery + "'";
  }
  
  try {
    const iterator = isFolder ? DriveApp.searchFolders(searchString) : DriveApp.searchFiles(searchString);
    let count = 0;
    while (iterator.hasNext() && count < 15) {
      const item = iterator.next();
      results.push({ id: item.getId(), name: item.getName() });
      count++;
    }
  } catch (e) {
    Logger.log("Error searching files: " + e.message);
  }
  return results;
}

/**
 * Mengekstrak tag berformat <<tag>> dari dokumen Google Doc
 */
function getDocTags(docId) {
  if (!docId) return [];
  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();
    const text = body.getText() || "";
    const regex = /<<([^>]+)>>/g;
    const tags = new Set();
    let match;
    
    // Cari di Body
    while ((match = regex.exec(text)) !== null) {
      tags.add(match[1].trim());
    }
    
    // Cari di Header
    const header = doc.getHeader();
    if (header) {
      const hText = header.getText() || "";
      const hRegex = /<<([^>]+)>>/g;
      while ((match = hRegex.exec(hText)) !== null) {
        tags.add(match[1].trim());
      }
    }
    
    // Cari di Footer
    const footer = doc.getFooter();
    if (footer) {
      const fText = footer.getText() || "";
      const fRegex = /<<([^>]+)>>/g;
      while ((match = fRegex.exec(fText)) !== null) {
        tags.add(match[1].trim());
      }
    }
    
    return Array.from(tags);
  } catch (e) {
    throw new Error("Gagal membaca tag dokumen: " + e.message);
  }
}

// ==========================================
// 4. PENYIMPANAN PENGATURAN PEKERJAAN (JOBS)
// ==========================================

function getSavedJobs() {
  const properties = PropertiesService.getDocumentProperties();
  const jobsData = properties.getProperty('MERGE_JOBS_WIZARD');
  return jobsData ? JSON.parse(jobsData) : {};
}

function saveJobSettings(jobName, settings) {
  if (!jobName) throw new Error("Nama pengaturan tidak boleh kosong.");
  const jobs = getSavedJobs();
  jobs[jobName] = settings;
  PropertiesService.getDocumentProperties().setProperty('MERGE_JOBS_WIZARD', JSON.stringify(jobs));
  return "Pengaturan '" + jobName + "' berhasil disimpan!";
}

function deleteJobSettings(jobName) {
  const jobs = getSavedJobs();
  if (jobs[jobName]) {
    delete jobs[jobName];
    PropertiesService.getDocumentProperties().setProperty('MERGE_JOBS_WIZARD', JSON.stringify(jobs));
    return "Pengaturan '" + jobName + "' berhasil dihapus!";
  }
  return "Pengaturan tidak ditemukan.";
}

// ==========================================
// 5. SISTEM POLLING PROGRESS (CACHE SERVICE)
// ==========================================
function getProgress() {
  const progress = CacheService.getUserCache().get('merge_progress_wizard');
  return progress ? JSON.parse(progress) : null;
}

function clearProgress() {
  CacheService.getUserCache().remove('merge_progress_wizard');
}

// ==========================================
// 6. PROSES UTAMA MERGE (DIJALANKAN DARI WIZARD UI)
// ==========================================
function runMergeProcess(jobName) {
  const jobs = getSavedJobs();
  const settings = jobs[jobName];
  if (!settings) throw new Error("Pengaturan '" + jobName + "' tidak ditemukan.");
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(settings.sheetName);
  if (!sheet) throw new Error("Sheet '" + settings.sheetName + "' tidak ditemukan.");
  
  const dataRange = sheet.getDataRange();
  const data = dataRange.getDisplayValues();
  if (data.length < 2) throw new Error("Sheet kosong atau hanya berisi baris header.");
  
  const headers = data[0].map(h => h.toString().trim());
  
  // Penamaan kolom status dinamis berbasis nama pekerjaan agar terisolasi
  const statusHeader = `[${jobName}] Status`;
  const linkHeader = `[${jobName}] Link`;
  
  let statusColIndex = headers.indexOf(statusHeader);
  let linkColIndex = headers.indexOf(linkHeader);
  
  // Jika kolom belum ada di sheet, tambahkan secara otomatis
  if (statusColIndex === -1) {
    statusColIndex = headers.length;
    sheet.getRange(1, statusColIndex + 1).setValue(statusHeader);
    headers.push(statusHeader);
  }
  if (linkColIndex === -1) {
    linkColIndex = headers.length;
    sheet.getRange(1, linkColIndex + 1).setValue(linkHeader);
    headers.push(linkHeader);
  }
  
  // Validasi dokumen template
  let templateFile;
  try {
    templateFile = DriveApp.getFileById(settings.templateDocId);
  } catch (e) {
    throw new Error("Template Google Doc tidak ditemukan. Periksa ID template.");
  }
  
  // Validasi folder penyimpanan tunggal jika diset static
  let singleFolder;
  if (settings.storageMode === 'single') {
    try {
      singleFolder = DriveApp.getFolderById(settings.destinationFolderId);
    } catch (e) {
      throw new Error("Folder penyimpanan tidak ditemukan. Periksa ID folder.");
    }
  }
  
  // Cari baris data yang lolos filter & belum diproses (untuk fitur resume)
  const rowsToProcess = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const currentStatus = row[statusColIndex];
    
    // Lewati jika sudah berstatus Success (lanjutkan kembali yang tertunda)
    if (currentStatus === 'Success' || (currentStatus && currentStatus.toString().startsWith('Success'))) {
      continue;
    }
    
    // Cek Filter Data
    if (settings.useFilter) {
      let keepRow = true;
      const filters = settings.filters || [];
      
      // Dukungan untuk pengaturan filter tunggal (legacy)
      if (filters.length === 0 && settings.filterColumn && settings.filterValue) {
        filters.push({
          column: settings.filterColumn,
          operator: 'equals',
          value: settings.filterValue
        });
      }
      
      for (let f = 0; f < filters.length; f++) {
        const filter = filters[f];
        const fColIdx = headers.indexOf(filter.column);
        if (fColIdx === -1) continue;
        
        const cellVal = row[fColIdx] !== null && row[fColIdx] !== undefined ? row[fColIdx].toString().trim() : '';
        const filterVal = filter.value ? filter.value.toString().trim() : '';
        const op = filter.operator || 'equals';
        
        let match = false;
        switch (op) {
          case 'equals':
            match = (cellVal.toLowerCase() === filterVal.toLowerCase());
            break;
          case 'notequals':
            match = (cellVal.toLowerCase() !== filterVal.toLowerCase());
            break;
          case 'contains':
            match = (cellVal.toLowerCase().indexOf(filterVal.toLowerCase()) !== -1);
            break;
          case 'notcontains':
            match = (cellVal.toLowerCase().indexOf(filterVal.toLowerCase()) === -1);
            break;
          case 'empty':
            match = (cellVal === '');
            break;
          case 'notempty':
            match = (cellVal !== '');
            break;
          default:
            match = true;
        }
        
        if (!match) {
          keepRow = false;
          break;
        }
      }
      
      if (!keepRow) {
        continue;
      }
    }
    
    rowsToProcess.push({
      rowIndex: i + 1, // Baris 1-indexed di sheet
      rowData: row
    });
  }
  
  const totalRows = rowsToProcess.length;
  if (totalRows === 0) {
    CacheService.getUserCache().put('merge_progress_wizard', JSON.stringify({
      current: 0,
      total: 0,
      status: 'Semua baris sudah diproses atau tidak ada data yang memenuhi filter.',
      success: 0,
      error: 0,
      completed: true
    }), 300);
    return { success: 0, error: 0, msg: "Tidak ada baris yang perlu diproses." };
  }
  
  let successCount = 0;
  let errorCount = 0;
  
  // Set progress awal
  CacheService.getUserCache().put('merge_progress_wizard', JSON.stringify({
    current: 0,
    total: totalRows,
    status: 'Menyiapkan berkas...',
    success: 0,
    error: 0,
    completed: false
  }), 300);
  
  // --- PENGATURAN MODE SINGLE DOCUMENT (SATU BERKAS GABUNGAN) ---
  let combinedDoc = null;
  let combinedBody = null;
  let combinedFile = null;
  let combinedFolder = null;
  
  if (settings.fileMode === 'single') {
    if (settings.storageMode === 'single') {
      combinedFolder = singleFolder;
    } else {
      // Dynamic folder: ambil lokasi folder dari baris pertama yang diproses
      const folderColIdx = headers.indexOf(settings.destinationFolderColumn);
      const folderId = folderColIdx !== -1 ? rowsToProcess[0].rowData[folderColIdx].toString().trim() : '';
      try {
        combinedFolder = DriveApp.getFolderById(folderId);
      } catch (e) {
        combinedFolder = DriveApp.getRootFolder();
      }
    }
    
    const combinedName = replaceFileNameTags(settings.fileNameFormat, headers, rowsToProcess[0].rowData) + " (Gabungan)";
    combinedFile = templateFile.makeCopy(combinedName, combinedFolder);
    combinedDoc = DocumentApp.openById(combinedFile.getId());
    combinedBody = combinedDoc.getBody();
    combinedBody.clear(); // Bersihkan isi awal sebelum digabung
  }
  
  // --- MULAI PERULANGAN PROSES MERGE ---
  for (let idx = 0; idx < rowsToProcess.length; idx++) {
    const item = rowsToProcess[idx];
    const rIndex = item.rowIndex;
    const rData = item.rowData;
    
    // Update progress
    CacheService.getUserCache().put('merge_progress_wizard', JSON.stringify({
      current: idx,
      total: totalRows,
      status: `Memproses baris ke-${rIndex} dari ${data.length}...`,
      success: successCount,
      error: errorCount,
      completed: false
    }), 300);
    
    const rowMap = {};
    headers.forEach((header, index) => {
      rowMap[header] = rData[index];
    });
    
    try {
      // Dapatkan folder penyimpanan baris ini
      let targetFolder = singleFolder;
      if (settings.storageMode === 'dynamic') {
        const folderColIdx = headers.indexOf(settings.destinationFolderColumn);
        const folderId = folderColIdx !== -1 ? rData[folderColIdx].toString().trim() : '';
        if (!folderId) throw new Error("ID folder dinamis kosong.");
        targetFolder = DriveApp.getFolderById(folderId);
      }
      
      const newFileName = replaceFileNameTags(settings.fileNameFormat, headers, rData);
      
      if (settings.fileMode === 'multiple') {
        // --- MULTIPLE DOKUMEN (BERKAS TERPISAH) ---
        const tempCopy = templateFile.makeCopy(newFileName, targetFolder);
        const copyDoc = DocumentApp.openById(tempCopy.getId());
        const copyBody = copyDoc.getBody();
        
        replaceTagsInDocElement(copyBody, settings.tagMappings, rowMap);
        
        const hSec = copyDoc.getHeader();
        if (hSec) replaceTagsInDocElement(hSec, settings.tagMappings, rowMap);
        const fSec = copyDoc.getFooter();
        if (fSec) replaceTagsInDocElement(fSec, settings.tagMappings, rowMap);
        
        copyDoc.saveAndClose();
        
        let finalFile = tempCopy;
        let fileLink = tempCopy.getUrl();
        
        if (settings.outputType === 'pdf') {
          const pdfBlob = tempCopy.getAs(MimeType.PDF);
          const pdfFile = targetFolder.createFile(pdfBlob);
          pdfFile.setName(newFileName + '.pdf');
          finalFile = pdfFile;
          fileLink = pdfFile.getUrl();
          tempCopy.setTrashed(true); // Hapus dokumen docx sementara
        }
        
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        sheet.getRange(rIndex, statusColIndex + 1).setValue('Success at ' + timestamp);
        sheet.getRange(rIndex, linkColIndex + 1).setValue(fileLink);
        
      } else {
        // --- SINGLE DOCUMENT (DIGABUNGKAN) ---
        // Salin template sementara untuk diisi nilainya untuk baris saat ini
        const tempCopy = templateFile.makeCopy("temp_merge_row_" + rIndex, combinedFolder);
        const tempDoc = DocumentApp.openById(tempCopy.getId());
        const tempBody = tempDoc.getBody();
        
        replaceTagsInDocElement(tempBody, settings.tagMappings, rowMap);
        
        const hSec = tempDoc.getHeader();
        if (hSec) replaceTagsInDocElement(hSec, settings.tagMappings, rowMap);
        const fSec = tempDoc.getFooter();
        if (fSec) replaceTagsInDocElement(fSec, settings.tagMappings, rowMap);
        
        tempDoc.saveAndClose();
        
        // Gabungkan body berkas sementara ke body berkas utama
        appendDocBodyToTarget(tempDoc.getId(), combinedBody);
        
        tempCopy.setTrashed(true); // Hapus sementara
        
        const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        sheet.getRange(rIndex, statusColIndex + 1).setValue('Merged at ' + timestamp);
      }
      
      successCount++;
    } catch (err) {
      Logger.log("Error pada baris ke-" + rIndex + ": " + err.message);
      sheet.getRange(rIndex, statusColIndex + 1).setValue('Error: ' + err.message);
      errorCount++;
    }
  }
  
  // Post-processing untuk single combined document
  if (settings.fileMode === 'single' && combinedDoc !== null) {
    combinedDoc.saveAndClose();
    let finalFile = combinedFile;
    let fileLink = combinedFile.getUrl();
    
    if (settings.outputType === 'pdf') {
      const pdfBlob = combinedFile.getAs(MimeType.PDF);
      const pdfFile = combinedFolder.createFile(pdfBlob);
      pdfFile.setName(combinedFile.getName() + '.pdf');
      finalFile = pdfFile;
      fileLink = pdfFile.getUrl();
      combinedFile.setTrashed(true); // Hapus salinan docx utama
    }
    
    // Tulis tautan berkas gabungan ke semua baris yang berhasil diproses
    rowsToProcess.forEach(item => {
      const statusVal = sheet.getRange(item.rowIndex, statusColIndex + 1).getValue().toString();
      if (statusVal.startsWith('Merged at')) {
        sheet.getRange(item.rowIndex, linkColIndex + 1).setValue(fileLink);
        sheet.getRange(item.rowIndex, statusColIndex + 1).setValue('Success ' + statusVal.substring(6));
      }
    });
  }
  
  // Simpan progress selesai
  CacheService.getUserCache().put('merge_progress_wizard', JSON.stringify({
    current: totalRows,
    total: totalRows,
    status: `Selesai! Sukses: ${successCount}, Gagal: ${errorCount}`,
    success: successCount,
    error: errorCount,
    completed: true
  }), 300);
  
  return {
    success: successCount,
    error: errorCount
  };
}

// ==========================================
// FUNGSI UTILITY (PEMBANTU MERGE)
// ==========================================

function replaceFileNameTags(formatStr, headers, rowData) {
  let result = formatStr;
  headers.forEach((header, index) => {
    const val = rowData[index] !== null && rowData[index] !== undefined ? rowData[index].toString() : '';
    const regex1 = new RegExp('<<' + escapeRegex(header) + '>>', 'g');
    const regex2 = new RegExp('{{' + escapeRegex(header) + '}}', 'g');
    result = result.replace(regex1, val).replace(regex2, val);
  });
  return result;
}

function isTrueValue(val) {
  if (!val) return false;
  const str = val.toString().trim().toLowerCase();
  return str === 'true' || str === '1' || str === 'yes' || str === 'y' || str === 'ya' || str === 'checked' || str === 'x' || str === 'benar';
}

function fetchImageBlob(urlOrId) {
  if (!urlOrId) return null;
  const cleanInput = urlOrId.toString().trim();
  if (!cleanInput) return null;
  
  let fileId = null;
  const driveUrlRegexes = [
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /id=([a-zA-Z0-9_-]+)/,
    /open\?id=([a-zA-Z0-9_-]+)/
  ];
  
  for (let i = 0; i < driveUrlRegexes.length; i++) {
    const match = cleanInput.match(driveUrlRegexes[i]);
    if (match && match[1]) {
      fileId = match[1];
      break;
    }
  }
  
  if (!fileId && !cleanInput.includes('/') && !cleanInput.includes('.') && cleanInput.length > 20) {
    fileId = cleanInput;
  }
  
  try {
    if (fileId) {
      const file = DriveApp.getFileById(fileId);
      return file.getBlob();
    } else if (cleanInput.startsWith('http://') || cleanInput.startsWith('https://')) {
      const response = UrlFetchApp.fetch(cleanInput);
      return response.getBlob();
    }
  } catch (e) {
    Logger.log("Gagal mengambil gambar dari: " + cleanInput + ". Error: " + e.message);
  }
  return null;
}

function replaceTextWithImage(container, searchText, imageUrl, widthCm, heightCm) {
  let found = container.findText(searchText);
  if (!found) return;
  
  const blob = fetchImageBlob(imageUrl);
  if (!blob) {
    container.replaceText(searchText, imageUrl || "");
    return;
  }
  
  while (found) {
    const textElement = found.getElement().asText();
    const startOffset = found.getStartOffset();
    const endOffset = found.getEndOffsetInclusive();
    const parent = textElement.getParent();
    
    let textToReplace = textElement.getText();
    let imageChildIndex;
    
    if (typeof parent.getChildIndex === 'function') {
      const childIndex = parent.getChildIndex(textElement);
      
      if (startOffset > 0 && endOffset < textToReplace.length - 1) {
        const prefix = textToReplace.substring(0, startOffset);
        const suffix = textToReplace.substring(endOffset + 1);
        
        textElement.setText(prefix);
        const suffixText = parent.insertText(childIndex + 1, suffix);
        const img = parent.insertInlineImage(childIndex + 1, blob);
        imageChildIndex = childIndex + 1;
        
      } else if (startOffset > 0) {
        const prefix = textToReplace.substring(0, startOffset);
        textElement.setText(prefix);
        const img = parent.insertInlineImage(childIndex + 1, blob);
        imageChildIndex = childIndex + 1;
        
      } else if (endOffset < textToReplace.length - 1) {
        const suffix = textToReplace.substring(endOffset + 1);
        textElement.setText(suffix);
        const img = parent.insertInlineImage(childIndex, blob);
        imageChildIndex = childIndex;
        
      } else {
        const img = parent.insertInlineImage(childIndex, blob);
        parent.removeChild(textElement);
        imageChildIndex = childIndex;
      }
      
      if (imageChildIndex !== undefined) {
        const img = parent.getChild(imageChildIndex);
        if (img && img.getType() === DocumentApp.ElementType.INLINE_IMAGE) {
          const cmToPoints = 72 / 2.54;
          if (widthCm && !isNaN(widthCm) && Number(widthCm) > 0) {
            img.setWidth(Number(widthCm) * cmToPoints);
          }
          if (heightCm && !isNaN(heightCm) && Number(heightCm) > 0) {
            img.setHeight(Number(heightCm) * cmToPoints);
          }
        }
      }
    } else {
      container.replaceText(searchText, imageUrl || "");
    }
    
    found = container.findText(searchText);
  }
}

function replaceTagsInDocElement(element, tagMappings, rowMap) {
  for (let tag in tagMappings) {
    const mapping = tagMappings[tag];
    let colHeader = '';
    let type = 'text';
    let width = '';
    let height = '';
    
    if (typeof mapping === 'string') {
      colHeader = mapping;
    } else if (mapping && typeof mapping === 'object') {
      colHeader = mapping.column || '';
      type = mapping.type || 'text';
      width = mapping.width || '';
      height = mapping.height || '';
    }
    
    if (!colHeader) {
      element.replaceText('<<' + escapeRegex(tag) + '>>', '');
      continue;
    }
    
    const val = (rowMap[colHeader] !== undefined && rowMap[colHeader] !== null) 
      ? rowMap[colHeader].toString() 
      : '';
      
    if (type === 'image') {
      replaceTextWithImage(element, '<<' + escapeRegex(tag) + '>>', val, width, height);
    } else if (type === 'checkbox') {
      const isTrue = isTrueValue(val);
      const checkboxChar = isTrue ? '☑' : '☐';
      element.replaceText('<<' + escapeRegex(tag) + '>>', checkboxChar);
    } else {
      element.replaceText('<<' + escapeRegex(tag) + '>>', val);
    }
  }
}

function escapeRegex(string) {
  return string.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
}

/**
 * Menyalin struktur dokumen dari satu Google Doc ke Google Doc tujuan
 */
function appendDocBodyToTarget(sourceDocId, targetBody) {
  const sourceDoc = DocumentApp.openById(sourceDocId);
  const sourceBody = sourceDoc.getBody();
  const numChildren = sourceBody.getNumChildren();
  
  // Sisipkan page break jika dokumen utama sudah memiliki isi
  if (targetBody.getNumChildren() > 0 && targetBody.getText().trim() !== '') {
    targetBody.appendPageBreak();
  }
  
  for (let i = 0; i < numChildren; i++) {
    const child = sourceBody.getChild(i).copy();
    const type = child.getType();
    
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      targetBody.appendParagraph(child);
    } else if (type === DocumentApp.ElementType.TABLE) {
      targetBody.appendTable(child);
    } else if (type === DocumentApp.ElementType.LIST_ITEM) {
      targetBody.appendListItem(child);
    } else if (type === DocumentApp.ElementType.INLINE_IMAGE) {
      const tempP = targetBody.appendParagraph("");
      tempP.appendInlineImage(child);
    } else {
      try {
        targetBody.appendParagraph(child.asText());
      } catch (err) {
        // Lewati tipe elemen yang tidak bisa disalin
      }
    }
  }
}

// ==========================================
// 7. KODE HTML SIDEBAR WIZARD (STRING TEMPLATE)
// ==========================================
function doGetSidebarHtml() {
  return `
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
  <style>
    :root {
      --primary: #4f46e5;
      --primary-hover: #4338ca;
      --bg-main: #f8fafc;
      --bg-card: #ffffff;
      --text-main: #0f172a;
      --text-muted: #64748b;
      --border: #e2e8f0;
      --focus-ring: rgba(79, 70, 229, 0.15);
      --success: #22c55e;
      --danger: #ef4444;
      --warning: #f59e0b;
    }

    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }

    body {
      font-family: 'Plus Jakarta Sans', sans-serif;
      background-color: var(--bg-main);
      color: var(--text-main);
      padding: 14px;
      font-size: 13px;
      line-height: 1.4;
      display: flex;
      flex-direction: column;
      height: 100vh;
      overflow: hidden;
    }

    /* Wizard Steps Header */
    .steps-indicator {
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 16px;
      background: white;
      padding: 8px 12px;
      border-radius: 10px;
      border: 1px solid var(--border);
    }

    .steps-indicator span {
      font-size: 11px;
      font-weight: 600;
      color: var(--text-muted);
    }

    .steps-indicator .step-active {
      color: var(--primary);
      font-weight: 700;
    }

    /* Content Area */
    .content-area {
      flex: 1;
      overflow-y: auto;
      margin-bottom: 60px;
      padding-right: 2px;
    }

    .wizard-step {
      display: none;
    }

    .wizard-step.active {
      display: block;
      animation: fadeIn 0.3s ease-in-out;
    }

    @keyframes fadeIn {
      from { opacity: 0; transform: translateY(4px); }
      to { opacity: 1; transform: translateY(0); }
    }

    h3 {
      font-size: 14px;
      font-weight: 700;
      color: var(--text-main);
      margin-bottom: 12px;
      border-bottom: 2px solid var(--border);
      padding-bottom: 6px;
    }

    .form-group {
      margin-bottom: 12px;
    }

    label {
      display: block;
      font-weight: 600;
      margin-bottom: 4px;
      color: var(--text-main);
      font-size: 11px;
    }

    .input-text, select, textarea {
      width: 100%;
      padding: 7px 10px;
      border: 1px solid var(--border);
      border-radius: 6px;
      font-size: 12px;
      font-family: inherit;
      color: var(--text-main);
      background-color: var(--bg-card);
      outline: none;
      transition: all 0.2s ease;
    }

    .input-text:focus, select:focus, textarea:focus {
      border-color: var(--primary);
      box-shadow: 0 0 0 2px var(--focus-ring);
    }

    .help-text {
      font-size: 10px;
      color: var(--text-muted);
      margin-top: 3px;
    }

    /* Search Results Box */
    .search-box {
      display: flex;
      gap: 6px;
    }

    .search-results {
      max-height: 120px;
      overflow-y: auto;
      border: 1px solid var(--border);
      border-radius: 6px;
      margin-top: 6px;
      background: white;
      display: none;
    }

    .search-item {
      padding: 6px 10px;
      cursor: pointer;
      border-bottom: 1px solid #f1f5f9;
      font-size: 11px;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .search-item:hover {
      background-color: #f1f5f9;
      color: var(--primary);
    }

    .selected-indicator {
      display: none;
      padding: 6px 10px;
      background-color: #e0e7ff;
      border: 1px solid #c7d2fe;
      border-radius: 6px;
      font-size: 11px;
      font-weight: 500;
      color: #3730a3;
      margin-top: 6px;
    }

    /* Tag Mapping Table */
    .mapping-table {
      margin-top: 10px;
      width: 100%;
      border-collapse: collapse;
    }

    .mapping-table th {
      text-align: left;
      font-size: 10px;
      font-weight: 700;
      color: var(--text-muted);
      padding-bottom: 4px;
      border-bottom: 1px solid var(--border);
    }

    .mapping-table td {
      padding: 5px 0;
      border-bottom: 1px dotted var(--border);
      vertical-align: middle;
    }

    .tag-name-col {
      font-family: monospace;
      font-size: 11px;
      color: #db2777;
      max-width: 110px;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    /* Navigation Footer */
    .nav-footer {
      position: absolute;
      bottom: 0;
      left: 0;
      right: 0;
      background-color: white;
      border-top: 1px solid var(--border);
      padding: 10px 14px;
      display: flex;
      gap: 8px;
    }

    .btn {
      flex: 1;
      padding: 9px 12px;
      border-radius: 6px;
      border: none;
      font-family: inherit;
      font-weight: 600;
      font-size: 12px;
      cursor: pointer;
      text-align: center;
      transition: all 0.2s ease;
    }

    .btn-next {
      background-color: var(--primary);
      color: white;
    }

    .btn-next:hover {
      background-color: var(--primary-hover);
    }

    .btn-back {
      background-color: transparent;
      color: var(--text-muted);
      border: 1px solid var(--border);
    }

    .btn-back:hover {
      background-color: #f1f5f9;
      color: var(--text-main);
    }

    .btn-search {
      padding: 0 12px;
      background-color: #f1f5f9;
      color: var(--text-main);
      border: 1px solid var(--border);
      cursor: pointer;
      border-radius: 6px;
    }

    .btn-search:hover {
      background-color: #e2e8f0;
    }

    /* Radio Group Styling */
    .radio-group {
      display: flex;
      gap: 12px;
      margin-top: 4px;
    }

    .radio-label {
      display: flex;
      align-items: center;
      gap: 4px;
      font-weight: 500;
      font-size: 11px;
      cursor: pointer;
    }

    /* Progress UI */
    .progress-card {
      background: white;
      border: 1px solid var(--border);
      border-radius: 8px;
      padding: 12px;
      margin-top: 12px;
      display: none;
    }

    .progress-header {
      display: flex;
      justify-content: space-between;
      margin-bottom: 6px;
      font-weight: 600;
    }

    .progress-bar-container {
      width: 100%;
      height: 5px;
      background-color: #f1f5f9;
      border-radius: 3px;
      overflow: hidden;
      margin-bottom: 8px;
    }

    .progress-bar-fill {
      height: 100%;
      width: 0%;
      background-color: var(--primary);
      transition: width 0.2s ease;
    }

    .progress-details {
      font-size: 11px;
      color: var(--text-muted);
    }

    .progress-stats {
      display: flex;
      justify-content: space-between;
      font-size: 10px;
      margin-top: 6px;
      border-top: 1px solid #f1f5f9;
      padding-top: 4px;
    }

    .stat-success { color: var(--success); font-weight: 600; }
    .stat-danger { color: var(--danger); font-weight: 600; }

    .disabled {
      opacity: 0.6;
      pointer-events: none;
    }
  </style>
</head>
<body>

  <!-- Indicators -->
  <div class="steps-indicator">
    <span id="indicatorStepText">Langkah 1 dari 5</span>
    <span class="step-active" id="indicatorTitle">Setup</span>
  </div>

  <!-- Content Area -->
  <div class="content-area" id="wizardContent">
    
    <!-- STEP 1: JOB SETUP & SHEET -->
    <div class="wizard-step active" id="step1">
      <h3>1. Setup Pekerjaan</h3>
      
      <div class="form-group">
        <label for="jobSelector">Pilih Profil Pengaturan</label>
        <select id="jobSelector" onchange="onJobChanged()">
          <option value="">-- Buat Pengaturan Baru --</option>
        </select>
      </div>

      <div class="form-group" id="newJobGroup">
        <label for="jobNameInput">Nama Pengaturan Baru</label>
        <input type="text" id="jobNameInput" class="input-text" placeholder="Misal: Cetak Invoice Penjualan">
      </div>

      <div class="form-group">
        <label for="sheetSelector">Pilih Sheet Aktif</label>
        <select id="sheetSelector" onchange="onSheetChanged()">
          <option value="">-- Memuat Sheet... --</option>
        </select>
      </div>
    </div>

    <!-- STEP 2: CHOOSE TEMPLATE & MAP TAGS -->
    <div class="wizard-step" id="step2">
      <h3>2. Pilih Template Dokumen</h3>
      
      <div class="form-group">
        <label for="docSearch">Cari Template Google Doc</label>
        <div class="search-box">
          <input type="text" id="docSearch" class="input-text" placeholder="Masukkan kata kunci...">
          <button type="button" class="btn-search" onclick="searchTemplates()">Cari</button>
        </div>
        
        <!-- Results -->
        <div class="search-results" id="docSearchResults"></div>
        <div class="selected-indicator" id="selectedDocIndicator"></div>
      </div>

      <!-- Tag Mappings -->
      <div id="tagMappingContainer" style="display:none; margin-top: 14px;">
        <label style="font-weight:700; display: block; margin-bottom: 6px;">Pemetaan Tag Template</label>
        <div id="tagMappingList">
          <!-- Dynamic Mappings -->
        </div>
      </div>
    </div>

    <!-- STEP 3: OUTPUT SETTINGS -->
    <div class="wizard-step" id="step3">
      <h3>3. Pengaturan Output</h3>
      
      <div class="form-group">
        <label>Format Berkas Output</label>
        <div class="radio-group">
          <label class="radio-label">
            <input type="radio" name="outputType" value="pdf" checked> PDF
          </label>
          <label class="radio-label">
            <input type="radio" name="outputType" value="doc"> Google Doc
          </label>
        </div>
      </div>

      <div class="form-group">
        <label>Model Pembuatan Berkas</label>
        <div class="radio-group">
          <label class="radio-label">
            <input type="radio" name="fileMode" value="multiple" checked> Berkas Terpisah (Multiple)
          </label>
          <label class="radio-label">
            <input type="radio" name="fileMode" value="single"> Gabung (Single File)
          </label>
        </div>
        <div class="help-text">Multiple: Satu baris satu file. Single: Semua baris digabung di satu file template.</div>
      </div>

      <div class="form-group">
        <label for="fileNameFormat">Format Nama File Output</label>
        <input type="text" id="fileNameFormat" class="input-text" placeholder="Misal: Invoice - <<Nama>>">
        <div class="help-text">Gunakan tag seperti <<Nama>> agar dinamis sesuai nilai kolom.</div>
      </div>
    </div>

    <!-- STEP 4: STORAGE LOCATION -->
    <div class="wizard-step" id="step4">
      <h3>4. Folder Penyimpanan</h3>
      
      <div class="form-group">
        <label>Metode Folder Tujuan</label>
        <div class="radio-group">
          <label class="radio-label">
            <input type="radio" name="storageMode" value="single" checked onchange="toggleStorageModeUI()"> Folder Tetap
          </label>
          <label class="radio-label">
            <input type="radio" name="storageMode" value="dynamic" onchange="toggleStorageModeUI()"> Dinamis per Baris
          </label>
        </div>
      </div>

      <!-- Static Folder UI -->
      <div class="form-group" id="staticStorageGroup">
        <label for="folderSearch">Cari Folder Google Drive</label>
        <div class="search-box">
          <input type="text" id="folderSearch" class="input-text" placeholder="Masukkan kata kunci...">
          <button type="button" class="btn-search" onclick="searchFolders()">Cari</button>
        </div>
        <div class="search-results" id="folderSearchResults"></div>
        <div class="selected-indicator" id="selectedFolderIndicator"></div>
      </div>

      <!-- Dynamic Folder UI -->
      <div class="form-group" id="dynamicStorageGroup" style="display:none;">
        <label for="folderColumnSelector">Kolom ID Folder Google Drive</label>
        <select id="folderColumnSelector">
          <!-- Populated from Sheet Headers -->
        </select>
        <div class="help-text">Pilih kolom yang berisi ID Folder Google Drive tujuan setiap baris.</div>
      </div>
    </div>

    <!-- STEP 5: FILTER & EXECUTION -->
    <div class="wizard-step" id="step5">
      <h3>5. Filter & Jalankan</h3>
      
      <div class="form-group">
        <label class="radio-label" style="font-weight:600;">
          <input type="checkbox" id="useFilterCheckbox" onchange="toggleFilterUI()"> Gunakan Filter Data
        </label>
      </div>

      <div id="filterInputsGroup" style="display:none; padding-left: 8px; border-left: 2px solid var(--border);">
        <div class="form-group" style="margin-bottom: 8px;">
          <label>Aturan Filter (AND)</label>
          <div id="filterRowsContainer" style="margin-bottom: 8px;">
            <!-- Dynamic filter rows -->
          </div>
          <button type="button" class="btn btn-back" style="padding: 4px 8px; font-size: 11px; margin-top: 6px; width: auto; display: inline-block;" onclick="addFilterRow()">+ Tambah Filter</button>
        </div>
      </div>

      <!-- Run Action -->
      <div style="margin-top: 20px;">
        <button type="button" class="btn btn-next" style="width: 100%; font-size: 13px;" id="btnRunProcess" onclick="startExecution()">Jalankan Pekerjaan</button>
        <button type="button" class="btn btn-back" style="width: 100%; margin-top: 6px;" id="btnDeleteJob" onclick="deleteCurrentJob()">Hapus Pengaturan Ini</button>
      </div>

      <!-- Progress Card -->
      <div class="progress-card" id="progressCard">
        <div class="progress-header">
          <span id="progressPercent">0%</span>
          <span id="progressTitle">Memproses...</span>
        </div>
        <div class="progress-bar-container">
          <div class="progress-bar-fill" id="progressBarFill"></div>
        </div>
        <div class="progress-details" id="progressDetails">Memproses data...</div>
        <div class="progress-stats">
          <span>Sukses: <span class="stat-success" id="progressSuccess">0</span></span>
          <span>Gagal: <span class="stat-danger" id="progressError">0</span></span>
        </div>
      </div>
    </div>

  </div>

  <!-- Navigation Buttons (Bottom) -->
  <div class="nav-footer" id="navFooter">
    <button type="button" class="btn btn-back" id="btnBack" onclick="changeStep(-1)" style="visibility:hidden;">Kembali</button>
    <button type="button" class="btn btn-next" id="btnNext" onclick="changeStep(1)">Selanjutnya</button>
  </div>

  <script>
    let currentStep = 1;
    const totalSteps = 5;
    
    function normalizeString(str) {
      if (!str) return '';
      return str.toString().toLowerCase().replace(/[^a-z0-9]/g, '').trim();
    }
    
    // Global Cache State
    let sheetHeaders = [];
    let savedJobs = {};
    let activeTemplateTags = [];
    
    // Temp State Selected Template & Folder
    let selectedTemplateId = '';
    let selectedTemplateName = '';
    let selectedFolderId = '';
    let selectedFolderName = '';
    
    let pollInterval;

    // Load Initial Data
    window.onload = function() {
      // 1. Get saved sheets
      google.script.run.withSuccessHandler(function(sheets) {
        const selector = document.getElementById('sheetSelector');
        selector.innerHTML = '<option value="">-- Pilih Sheet --</option>';
        sheets.forEach(name => {
          selector.innerHTML += '<option value="' + name + '">' + name + '</option>';
        });
        
        // 2. Get saved jobs after sheets are loaded
        loadSavedJobsList();
      }).getSpreadsheetSheets();
    };

    function loadSavedJobsList() {
      google.script.run.withSuccessHandler(function(jobs) {
        savedJobs = jobs;
        const selector = document.getElementById('jobSelector');
        
        // Keep "Buat Baru"
        selector.innerHTML = '<option value="">-- Buat Pengaturan Baru --</option>';
        for (let name in jobs) {
          selector.innerHTML += '<option value="' + name + '">' + name + '</option>';
        }
      }).getSavedJobs();
    }

    // Handlers Step 1
    function onJobChanged() {
      const selectedJob = document.getElementById('jobSelector').value;
      const newJobGroup = document.getElementById('newJobGroup');
      const jobNameInput = document.getElementById('jobNameInput');
      
      if (selectedJob === '') {
        newJobGroup.style.display = 'block';
        jobNameInput.value = '';
        resetWizardState();
      } else {
        newJobGroup.style.display = 'none';
        loadJobSettingsIntoUI(selectedJob);
      }
    }

    function onSheetChanged() {
      const sheetName = document.getElementById('sheetSelector').value;
      if (!sheetName) return;
      
      // Load Sheet Headers
      google.script.run.withSuccessHandler(function(headers) {
        sheetHeaders = headers;
        
        // Update all header selectors
        updateHeaderDropdowns();
      }).getSheetHeaders(sheetName);
    }

    function updateHeaderDropdowns() {
      // Update folder dropdown selector (Step 4)
      const folderColSel = document.getElementById('folderColumnSelector');
      folderColSel.innerHTML = '<option value="">-- Pilih Kolom Folder ID --</option>';
      sheetHeaders.forEach(h => {
        folderColSel.innerHTML += '<option value="' + h + '">' + h + '</option>';
      });
      
      // Update all filter-col-selectors in existing filter rows
      const colSelectors = document.querySelectorAll('.filter-col-selector');
      colSelectors.forEach(select => {
        const currentVal = select.value;
        select.innerHTML = '<option value="">-- Pilih Kolom --</option>';
        sheetHeaders.forEach(h => {
          const selected = h === currentVal ? 'selected' : '';
          select.innerHTML += '<option value="' + h + '" ' + selected + '>' + h + '</option>';
        });
      });
      
      // Re-populate mappings if tags are loaded
      if (activeTemplateTags.length > 0) {
        renderTagMappingsTable();
      }
    }

    // Step 2 Template Search
    function searchTemplates() {
      const query = document.getElementById('docSearch').value;
      const resultsContainer = document.getElementById('docSearchResults');
      resultsContainer.style.display = 'block';
      resultsContainer.innerHTML = '<div style="padding: 10px; font-size:11px;">Mencari dokumen...</div>';
      
      google.script.run
        .withSuccessHandler(function(results) {
          resultsContainer.innerHTML = '';
          if (results.length === 0) {
            resultsContainer.innerHTML = '<div style="padding: 10px; font-size:11px; color:var(--text-muted);">Tidak ada file ditemukan.</div>';
            return;
          }
          results.forEach(file => {
            const div = document.createElement('div');
            div.className = 'search-item';
            div.innerText = file.name;
            div.onclick = function() {
              selectTemplate(file.id, file.name);
              resultsContainer.style.display = 'none';
            };
            resultsContainer.appendChild(div);
          });
        })
        .withFailureHandler(function(err) {
          resultsContainer.innerHTML = '<div style="padding: 10px; font-size:11px; color:var(--danger);">Error: ' + err.message + '</div>';
        })
        .searchDriveFiles(query, false);
    }

    function selectTemplate(id, name) {
      selectedTemplateId = id;
      selectedTemplateName = name;
      
      const indicator = document.getElementById('selectedDocIndicator');
      indicator.style.display = 'block';
      indicator.innerText = 'Terpilih: ' + name;
      
      // Load tags
      const mappingContainer = document.getElementById('tagMappingContainer');
      mappingContainer.style.display = 'block';
      document.getElementById('tagMappingList').innerHTML = '<div style="text-align:center; padding: 10px; font-size: 11px;">Membaca tag template...</div>';
      
      google.script.run.withSuccessHandler(function(tags) {
        activeTemplateTags = tags;
        renderTagMappingsTable();
      }).getDocTags(id);
    }

    function renderTagMappingsTable() {
      const container = document.getElementById('tagMappingList');
      container.innerHTML = '';
      
      if (activeTemplateTags.length === 0) {
        container.innerHTML = '<div style="text-align:center; padding: 10px; font-size: 11px; color:var(--warning);">Tidak ada tag &lt;&lt;...&gt;&gt; ditemukan pada template.</div>';
        return;
      }
      
      activeTemplateTags.forEach(tag => {
        const item = document.createElement('div');
        item.className = 'tag-mapping-item';
        item.style.borderBottom = '1px solid var(--border)';
        item.style.padding = '8px 0';
        
        // Row 1: Tag Name
        const headerRow = document.createElement('div');
        headerRow.style.display = 'flex';
        headerRow.style.justifyContent = 'space-between';
        headerRow.style.marginBottom = '4px';
        
        const label = document.createElement('span');
        label.style.fontFamily = 'monospace';
        label.style.fontSize = '11px';
        label.style.color = '#db2777';
        label.style.fontWeight = '600';
        label.innerText = '<<' + tag + '>>';
        headerRow.appendChild(label);
        item.appendChild(headerRow);
        
        // Row 2: Mappings control inputs
        const controlRow = document.createElement('div');
        controlRow.style.display = 'flex';
        controlRow.style.gap = '6px';
        
        const selectCol = document.createElement('select');
        selectCol.className = 'mapping-select';
        selectCol.style.flex = '1';
        selectCol.style.padding = '4px';
        selectCol.style.fontSize = '11px';
        selectCol.setAttribute('data-tag', tag);
        
        selectCol.innerHTML = '<option value="">-- Jangan Petakan --</option>';
        sheetHeaders.forEach(h => {
          const selectedAttr = (normalizeString(tag) === normalizeString(h)) ? 'selected' : '';
          selectCol.innerHTML += '<option value="' + h + '" ' + selectedAttr + '>' + h + '</option>';
        });
        controlRow.appendChild(selectCol);
        
        const selectType = document.createElement('select');
        selectType.className = 'type-select';
        selectType.style.width = '90px';
        selectType.style.padding = '4px';
        selectType.style.fontSize = '11px';
        selectType.setAttribute('data-tag', tag);
        selectType.onchange = function() {
          onTagTypeChanged(tag);
        };
        
        selectType.innerHTML = 
          '<option value="text" selected>Text</option>' +
          '<option value="image">Image</option>' +
          '<option value="checkbox">Check Box</option>';
        controlRow.appendChild(selectType);
        item.appendChild(controlRow);
        
        // Row 3: Image Options (width & height in cm)
        const imgOpts = document.createElement('div');
        imgOpts.className = 'image-options-container';
        imgOpts.id = 'image-options-' + tag;
        imgOpts.style.display = 'none';
        imgOpts.style.background = '#f8fafc';
        imgOpts.style.border = '1px solid var(--border)';
        imgOpts.style.borderRadius = '4px';
        imgOpts.style.padding = '6px';
        imgOpts.style.marginTop = '4px';
        
        imgOpts.innerHTML = 
          '<div style="display: flex; align-items: center; gap: 4px;">' +
            '<span style="font-size: 10px; color: var(--text-muted);">Lebar:</span>' +
            '<input type="number" step="any" min="0.1" class="image-width input-text" data-tag="' + tag + '" style="width: 50px; padding: 2px 4px; font-size: 10px; display: inline-block;" placeholder="cm">' +
            '<span style="font-size: 10px; color: var(--text-muted); margin-right: 6px;">cm</span>' +
            '<span style="font-size: 10px; color: var(--text-muted);">Tinggi:</span>' +
            '<input type="number" step="any" min="0.1" class="image-height input-text" data-tag="' + tag + '" style="width: 50px; padding: 2px 4px; font-size: 10px; display: inline-block;" placeholder="cm">' +
            '<span style="font-size: 10px; color: var(--text-muted);">cm</span>' +
          '</div>';
        item.appendChild(imgOpts);
        
        container.appendChild(item);
      });
    }
    
    function onTagTypeChanged(tag) {
      const typeSelect = document.querySelector('.type-select[data-tag="' + tag + '"]');
      const imgOpts = document.getElementById('image-options-' + tag);
      if (typeSelect && imgOpts) {
        if (typeSelect.value === 'image') {
          imgOpts.style.display = 'block';
        } else {
          imgOpts.style.display = 'none';
        }
      }
    }

    // Step 4 Storage Search
    function toggleStorageModeUI() {
      const mode = document.querySelector('input[name="storageMode"]:checked').value;
      const staticGroup = document.getElementById('staticStorageGroup');
      const dynamicGroup = document.getElementById('dynamicStorageGroup');
      
      if (mode === 'single') {
        staticGroup.style.display = 'block';
        dynamicGroup.style.display = 'none';
      } else {
        staticGroup.style.display = 'none';
        dynamicGroup.style.display = 'block';
      }
    }

    function searchFolders() {
      const query = document.getElementById('folderSearch').value;
      const resultsContainer = document.getElementById('folderSearchResults');
      resultsContainer.style.display = 'block';
      resultsContainer.innerHTML = '<div style="padding: 10px; font-size:11px;">Mencari folder...</div>';
      
      google.script.run
        .withSuccessHandler(function(results) {
          resultsContainer.innerHTML = '';
          if (results.length === 0) {
            resultsContainer.innerHTML = '<div style="padding: 10px; font-size:11px; color:var(--text-muted);">Tidak ada folder ditemukan.</div>';
            return;
          }
          results.forEach(folder => {
            const div = document.createElement('div');
            div.className = 'search-item';
            div.innerText = folder.name;
            div.onclick = function() {
              selectFolder(folder.id, folder.name);
              resultsContainer.style.display = 'none';
            };
            resultsContainer.appendChild(div);
          });
        })
        .withFailureHandler(function(err) {
          resultsContainer.innerHTML = '<div style="padding: 10px; font-size:11px; color:var(--danger);">Error: ' + err.message + '</div>';
        })
        .searchDriveFiles(query, true);
    }

    function selectFolder(id, name) {
      selectedFolderId = id;
      selectedFolderName = name;
      const indicator = document.getElementById('selectedFolderIndicator');
      indicator.style.display = 'block';
      indicator.innerText = 'Terpilih: ' + name;
    }

    // Step 5 Filter
    function toggleFilterUI() {
      const checked = document.getElementById('useFilterCheckbox').checked;
      document.getElementById('filterInputsGroup').style.display = checked ? 'block' : 'none';
    }

    function addFilterRow(savedCol, savedOp, savedVal) {
      const container = document.getElementById('filterRowsContainer');
      const rowDiv = document.createElement('div');
      rowDiv.className = 'filter-row';
      rowDiv.style.display = 'flex';
      rowDiv.style.flexDirection = 'column';
      rowDiv.style.gap = '4px';
      rowDiv.style.padding = '8px';
      rowDiv.style.marginBottom = '8px';
      rowDiv.style.border = '1px solid var(--border)';
      rowDiv.style.borderRadius = '6px';
      rowDiv.style.backgroundColor = '#f8fafc';
      
      // Column Select
      const colSel = document.createElement('select');
      colSel.className = 'filter-col-selector';
      colSel.style.fontSize = '11px';
      colSel.style.padding = '4px';
      colSel.innerHTML = '<option value="">-- Pilih Kolom --</option>';
      sheetHeaders.forEach(h => {
        const selected = h === savedCol ? 'selected' : '';
        colSel.innerHTML += '<option value="' + h + '" ' + selected + '>' + h + '</option>';
      });
      
      // Operator Select
      const opSel = document.createElement('select');
      opSel.className = 'filter-op-selector';
      opSel.style.fontSize = '11px';
      opSel.style.padding = '4px';
      opSel.innerHTML = 
        '<option value="equals">Sama Dengan</option>' +
        '<option value="notequals">Tidak Sama Dengan</option>' +
        '<option value="contains">Mengandung</option>' +
        '<option value="notcontains">Tidak Mengandung</option>' +
        '<option value="empty">Kosong</option>' +
        '<option value="notempty">Tidak Kosong</option>';
      if (savedOp) opSel.value = savedOp;
      
      // Value Input
      const valInput = document.createElement('input');
      valInput.type = 'text';
      valInput.className = 'filter-val-input input-text';
      valInput.style.fontSize = '11px';
      valInput.style.padding = '4px 8px';
      valInput.placeholder = 'Nilai filter...';
      if (savedVal) valInput.value = savedVal;
      
      // Show/hide valInput based on operator selection
      opSel.onchange = function() {
        if (opSel.value === 'empty' || opSel.value === 'notempty') {
          valInput.style.display = 'none';
        } else {
          valInput.style.display = 'block';
        }
      };
      // Trigger initially
      if (savedOp === 'empty' || savedOp === 'notempty') {
        valInput.style.display = 'none';
      }
      
      // Delete Button row
      const actionRow = document.createElement('div');
      actionRow.style.display = 'flex';
      actionRow.style.justifyContent = 'flex-end';
      
      const delBtn = document.createElement('button');
      delBtn.type = 'button';
      delBtn.className = 'btn btn-back';
      delBtn.style.padding = '2px 6px';
      delBtn.style.fontSize = '10px';
      delBtn.style.color = 'var(--danger)';
      delBtn.style.borderColor = 'var(--danger)';
      delBtn.style.width = 'auto';
      delBtn.style.marginTop = '2px';
      delBtn.innerText = 'Hapus';
      delBtn.onclick = function() {
        rowDiv.remove();
      };
      
      actionRow.appendChild(delBtn);
      
      rowDiv.appendChild(colSel);
      rowDiv.appendChild(opSel);
      rowDiv.appendChild(valInput);
      rowDiv.appendChild(actionRow);
      
      container.appendChild(rowDiv);
    }

    // Wizard Navigation Controller
    function changeStep(val) {
      // Validasi sebelum lanjut ke langkah berikutnya
      if (val > 0 && !validateStep(currentStep)) return;
      
      const newStep = currentStep + val;
      if (newStep < 1 || newStep > totalSteps) return;
      
      // Hide current, show new
      document.getElementById('step' + currentStep).classList.remove('active');
      document.getElementById('step' + newStep).classList.add('active');
      
      currentStep = newStep;
      
      // Update Navigation Buttons visibility
      document.getElementById('btnBack').style.visibility = (currentStep === 1) ? 'hidden' : 'visible';
      
      if (currentStep === totalSteps) {
        document.getElementById('btnNext').style.display = 'none';
      } else {
        document.getElementById('btnNext').style.display = 'block';
      }
      
      // Update Step Indicators
      document.getElementById('indicatorStepText').innerText = 'Langkah ' + currentStep + ' dari ' + totalSteps;
      
      const titles = ['Setup Pekerjaan', 'Pilih Template', 'Pengaturan Output', 'Lokasi Simpan', 'Jalankan'];
      document.getElementById('indicatorTitle').innerText = titles[currentStep - 1];
    }

    function validateStep(step) {
      if (step === 1) {
        const isNewJob = document.getElementById('jobSelector').value === '';
        const jobName = isNewJob ? document.getElementById('jobNameInput').value.trim() : document.getElementById('jobSelector').value;
        const sheetName = document.getElementById('sheetSelector').value;
        
        if (!jobName) {
          alert('Nama pengaturan tidak boleh kosong.');
          return false;
        }
        if (!sheetName) {
          alert('Silakan pilih sheet yang akan digunakan.');
          return false;
        }
      }
      if (step === 2) {
        if (!selectedTemplateId) {
          alert('Silakan cari dan pilih berkas template terlebih dahulu.');
          return false;
        }
      }
      if (step === 3) {
        const fileNameFormat = document.getElementById('fileNameFormat').value.trim();
        if (!fileNameFormat) {
          alert('Format nama berkas output tidak boleh kosong.');
          return false;
        }
      }
      if (step === 4) {
        const storageMode = document.querySelector('input[name="storageMode"]:checked').value;
        if (storageMode === 'single' && !selectedFolderId) {
          alert('Silakan cari dan pilih folder penyimpanan Google Drive.');
          return false;
        }
        if (storageMode === 'dynamic' && !document.getElementById('folderColumnSelector').value) {
          alert('Silakan pilih kolom ID folder dinamis.');
          return false;
        }
      }
      return true;
    }

    // Load Settings
    function loadJobSettingsIntoUI(jobName) {
      const settings = savedJobs[jobName];
      if (!settings) return;
      
      // Step 1
      document.getElementById('sheetSelector').value = settings.sheetName;
      onSheetChanged(); // trigger load headers
      
      // Step 2
      selectedTemplateId = settings.templateDocId;
      selectedTemplateName = settings.templateDocName;
      const docIndicator = document.getElementById('selectedDocIndicator');
      docIndicator.style.display = 'block';
      docIndicator.innerText = 'Terpilih: ' + settings.templateDocName;
      
      // Load tags
      document.getElementById('tagMappingContainer').style.display = 'block';
      document.getElementById('tagMappingList').innerHTML = '<div style="text-align:center; padding: 10px; font-size: 11px;">Membaca tag template...</div>';
      google.script.run.withSuccessHandler(function(tags) {
        activeTemplateTags = tags;
        renderTagMappingsTable();
        
        // Wait and select values mapping
        setTimeout(function() {
          const selects = document.querySelectorAll('.mapping-select');
          selects.forEach(select => {
            const tag = select.getAttribute('data-tag');
            const mapping = settings.tagMappings[tag];
            if (mapping) {
              let column = '';
              let type = 'text';
              let width = '';
              let height = '';
              
              if (typeof mapping === 'string') {
                column = mapping;
              } else if (typeof mapping === 'object') {
                column = mapping.column || '';
                type = mapping.type || 'text';
                width = mapping.width || '';
                height = mapping.height || '';
              }
              
              select.value = column;
              
              const typeSelect = document.querySelector('.type-select[data-tag="' + tag + '"]');
              if (typeSelect) {
                typeSelect.value = type;
                onTagTypeChanged(tag);
              }
              
              const widthInput = document.querySelector('.image-width[data-tag="' + tag + '"]');
              if (widthInput) widthInput.value = width;
              
              const heightInput = document.querySelector('.image-height[data-tag="' + tag + '"]');
              if (heightInput) heightInput.value = height;
            }
          });
        }, 800);
      }).getDocTags(settings.templateDocId);
      
      // Step 3
      document.querySelector('input[name="outputType"][value="' + settings.outputType + '"]').checked = true;
      document.querySelector('input[name="fileMode"][value="' + settings.fileMode + '"]').checked = true;
      document.getElementById('fileNameFormat').value = settings.fileNameFormat;
      
      // Step 4
      document.querySelector('input[name="storageMode"][value="' + settings.storageMode + '"]').checked = true;
      toggleStorageModeUI();
      if (settings.storageMode === 'single') {
        selectedFolderId = settings.destinationFolderId;
        selectedFolderName = settings.destinationFolderName;
        const fIndicator = document.getElementById('selectedFolderIndicator');
        fIndicator.style.display = 'block';
        fIndicator.innerText = 'Terpilih: ' + settings.destinationFolderName;
      } else {
        setTimeout(function() {
          document.getElementById('folderColumnSelector').value = settings.destinationFolderColumn;
        }, 800);
      }
      
      // Step 5
      document.getElementById('useFilterCheckbox').checked = settings.useFilter;
      toggleFilterUI();
      document.getElementById('filterRowsContainer').innerHTML = '';
      if (settings.useFilter) {
        setTimeout(function() {
          if (settings.filters && settings.filters.length > 0) {
            settings.filters.forEach(f => {
              addFilterRow(f.column, f.operator, f.value);
            });
          } else if (settings.filterColumn && settings.filterValue) {
            // Konversi dari pengaturan filter legacy
            addFilterRow(settings.filterColumn, 'equals', settings.filterValue);
          }
        }, 800);
      }
    }

    function resetWizardState() {
      document.getElementById('sheetSelector').value = '';
      selectedTemplateId = '';
      selectedTemplateName = '';
      document.getElementById('selectedDocIndicator').style.display = 'none';
      document.getElementById('tagMappingContainer').style.display = 'none';
      document.getElementById('tagMappingList').innerHTML = '';
      activeTemplateTags = [];
      
      document.querySelector('input[name="outputType"][value="pdf"]').checked = true;
      document.querySelector('input[name="fileMode"][value="multiple"]').checked = true;
      document.getElementById('fileNameFormat').value = '';
      
      document.querySelector('input[name="storageMode"][value="single"]').checked = true;
      toggleStorageModeUI();
      selectedFolderId = '';
      selectedFolderName = '';
      document.getElementById('selectedFolderIndicator').style.display = 'none';
      
      document.getElementById('useFilterCheckbox').checked = false;
      toggleFilterUI();
      document.getElementById('filterRowsContainer').innerHTML = '';
    }

    // Save and Execution Controls
    function getWizardSettings() {
      // Parse Mappings
      const tagMappings = {};
      const selects = document.querySelectorAll('.mapping-select');
      selects.forEach(select => {
        const tag = select.getAttribute('data-tag');
        const column = select.value;
        
        // Find corresponding type select
        const typeSelect = document.querySelector('.type-select[data-tag="' + tag + '"]');
        const type = typeSelect ? typeSelect.value : 'text';
        
        // Find width and height inputs
        const widthInput = document.querySelector('.image-width[data-tag="' + tag + '"]');
        const heightInput = document.querySelector('.image-height[data-tag="' + tag + '"]');
        
        const width = widthInput ? widthInput.value : '';
        const height = heightInput ? heightInput.value : '';
        
        tagMappings[tag] = {
          column: column,
          type: type,
          width: width,
          height: height
        };
      });

      const isNewJob = document.getElementById('jobSelector').value === '';
      const jobName = isNewJob ? document.getElementById('jobNameInput').value.trim() : document.getElementById('jobSelector').value;

      const filters = [];
      const filterRows = document.querySelectorAll('.filter-row');
      filterRows.forEach(row => {
        const col = row.querySelector('.filter-col-selector').value;
        const op = row.querySelector('.filter-op-selector').value;
        const val = row.querySelector('.filter-val-input').value;
        if (col) {
          filters.push({
            column: col,
            operator: op,
            value: val
          });
        }
      });

      return {
        jobName: jobName,
        sheetName: document.getElementById('sheetSelector').value,
        templateDocId: selectedTemplateId,
        templateDocName: selectedTemplateName,
        tagMappings: tagMappings,
        outputType: document.querySelector('input[name="outputType"]:checked').value,
        fileMode: document.querySelector('input[name="fileMode"]:checked').value,
        fileNameFormat: document.getElementById('fileNameFormat').value.trim(),
        storageMode: document.querySelector('input[name="storageMode"]:checked').value,
        destinationFolderId: selectedFolderId,
        destinationFolderName: selectedFolderName,
        destinationFolderColumn: document.getElementById('folderColumnSelector').value,
        useFilter: document.getElementById('useFilterCheckbox').checked,
        filters: filters
      };
    }

    function startExecution() {
      const settings = getWizardSettings();
      const jobName = settings.jobName;
      
      // Save settings first
      document.getElementById('btnRunProcess').disabled = true;
      document.getElementById('btnRunProcess').innerText = 'Menyimpan & Memulai...';
      
      google.script.run.withSuccessHandler(function() {
        // Run process
        document.getElementById('progressCard').style.display = 'block';
        resetProgressUI();
        
        // Start script execution
        google.script.run
          .withSuccessHandler(function(result) {
            stopPolling();
            alert('Merge Dokumen Selesai!\\nSukses: ' + result.success + '\\nGagal: ' + result.error);
            document.getElementById('btnRunProcess').disabled = false;
            document.getElementById('btnRunProcess').innerText = 'Jalankan Pekerjaan';
            google.script.run.clearProgress();
            
            // Reload list
            loadSavedJobsList();
          })
          .withFailureHandler(function(err) {
            stopPolling();
            alert('Error saat mengeksekusi merge: ' + err.message);
            document.getElementById('btnRunProcess').disabled = false;
            document.getElementById('btnRunProcess').innerText = 'Jalankan Pekerjaan';
            document.getElementById('progressTitle').innerText = 'Gagal!';
            document.getElementById('progressDetails').innerText = err.message;
            google.script.run.clearProgress();
          })
          .runMergeProcess(jobName);
          
        startPolling();
      }).saveJobSettings(jobName, settings);
    }

    function deleteCurrentJob() {
      const selectedJob = document.getElementById('jobSelector').value;
      if (!selectedJob) {
        alert('Pilih profil pengaturan yang valid terlebih dahulu.');
        return;
      }
      if (confirm('Apakah Anda yakin ingin menghapus profil pengaturan "' + selectedJob + '"?')) {
        google.script.run.withSuccessHandler(function(res) {
          alert(res);
          resetWizardState();
          loadSavedJobsList();
          document.getElementById('jobSelector').value = '';
          onJobChanged();
        }).deleteJobSettings(selectedJob);
      }
    }

    // Polling Progress
    function resetProgressUI() {
      document.getElementById('progressPercent').innerText = '0%';
      document.getElementById('progressBarFill').style.width = '0%';
      document.getElementById('progressTitle').innerText = 'Menghubungkan...';
      document.getElementById('progressDetails').innerText = 'Menganalisis baris...';
      document.getElementById('progressSuccess').innerText = '0';
      document.getElementById('progressError').innerText = '0';
    }

    function startPolling() {
      pollInterval = setInterval(function() {
        google.script.run.withSuccessHandler(function(progress) {
          if (progress) {
            const pct = progress.total > 0 ? Math.round((progress.current / progress.total) * 100) : 0;
            document.getElementById('progressPercent').innerText = pct + '%';
            document.getElementById('progressBarFill').style.width = pct + '%';
            document.getElementById('progressTitle').innerText = progress.completed ? 'Selesai!' : 'Memproses...';
            document.getElementById('progressDetails').innerText = progress.status;
            document.getElementById('progressSuccess').innerText = progress.success;
            document.getElementById('progressError').innerText = progress.error;
            
            if (progress.completed) {
              stopPolling();
            }
          }
        }).getProgress();
      }, 1500);
    }

    function stopPolling() {
      if (pollInterval) {
        clearInterval(pollInterval);
      }
    }
  </script>
</body>
</html>
  `;
}

