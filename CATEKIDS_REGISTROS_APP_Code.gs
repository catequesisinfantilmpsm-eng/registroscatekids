/**
 * CATEKIDS REGISTROS APP
 * Sincronización segura entre la app y los cuatro libros de registro.
 *
 * INSTALACIÓN
 * 1. Reemplaza todo el contenido de Code.gs por este archivo.
 * 2. Implementar > Administrar implementaciones > Editar.
 * 3. Ejecutar como: Yo. Acceso: Cualquier usuario.
 * 4. Crea una versión nueva y pulsa Implementar.
 *
 * Este proyecto es independiente. No modifica Apps Script de otras apps.
 */

const API_TOKEN = 'CKR_2026_MISERICORDIA_7f4a91';

const BOOKS = {
  'INICIACIÓN': {
    id: '1unbI6q36wnBb5B_CFpb5H67EdXtDEiGBFgFSycjgPK8',
    sheets: ['SILVIA','MARLEN','SAMANTHA','AURORA','TADEO','ANA','EDIL','JOHANA','MILDRED','LILIANA']
  },
  'CONFIRMACIÓN': {
    id: '1ob22xfpCAK0FiV6DUP77RNuTPr4a2LvI4qnxtZMX3JI',
    sheets: ['VICKY','HILIAMOR','RAQUEL','CECILIA ','ESTHER','FIDE DE JESUS','DON CHUY','CONCHITA','ARACELI','ESTRELLA']
  },
  'COMUNIÓN': {
    id: '1oTXcfavomdW_wIBabTdEIZME2XuGZo1heQj1i2Q3dUY',
    sheets: ['MANDY','TINA','ROX','MILDRED','PATY','JORGE','MONSE','CONCHITA','ALIZA','RICARDA','CAYETANA','LIDIA']
  },
  'PERSEVERANTES': {
    id: '1hPTqr2zn4FpC8szHPBORdjx35p1572Yb5y8cPgLo7as',
    sheets: ['HILIAMOR PERS','FIDENCIO','EDIL PERS']
  }
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('CATEKIDS · Registros')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

function getConfig() { return publicConfig_(); }
function getStudents(course) { return listStudents_(course); }
function saveStudent(student) { return upsertStudent_(student || {}); }

function uploadPdf(payload) {
  if (!payload || !payload.base64 || !payload.name || !payload.code) throw new Error('Archivo incompleto');
  const folders = DriveApp.getFoldersByName('CATEKIDS REGISTROS PDF');
  const root = folders.hasNext() ? folders.next() : DriveApp.createFolder('CATEKIDS REGISTROS PDF');
  const safeCode = String(payload.code).replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ _-]/g,'').trim();
  const matches = root.getFoldersByName(safeCode);
  const folder = matches.hasNext() ? matches.next() : root.createFolder(safeCode);
  const bytes = Utilities.base64Decode(String(payload.base64).split(',').pop());
  const blob = Utilities.newBlob(bytes,'application/pdf',payload.name);
  const file = folder.createFile(blob);
  return {id:file.getId(),name:file.getName(),url:file.getUrl()};
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.token !== API_TOKEN) return json_({ok:false,error:'Acceso no autorizado'});

    if (body.action === 'list') return json_({ok:true,students:listStudents_(body.course)});
    if (body.action === 'upsert') {
      lock.waitLock(20000);
      const result = upsertStudent_(body.student || {});
      log_('UPSERT', result.course, result.catechist, result.code, 'OK');
      return json_({ok:true,student:result});
    }
    if (body.action === 'config') return json_({ok:true,courses:publicConfig_()});
    return json_({ok:false,error:'Acción no reconocida'});
  } catch (err) {
    log_('ERROR','','','',String(err && err.message || err));
    return json_({ok:false,error:String(err && err.message || err)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function listStudents_(onlyCourse) {
  const output = [];
  Object.keys(BOOKS).forEach(course => {
    if (onlyCourse && normalize_(onlyCourse) !== normalize_(course)) return;
    const book = BOOKS[course];
    const ss = SpreadsheetApp.openById(book.id);
    book.sheets.forEach(sheetName => {
      const sh = ss.getSheetByName(sheetName);
      if (!sh || sh.getLastRow() < 3) return;
      const lastCol = Math.max(1, sh.getLastColumn());
      const headers = sh.getRange(2,1,1,lastCol).getDisplayValues()[0];
      const values = sh.getRange(3,1,sh.getLastRow()-2,lastCol).getValues();
      const idx = headerIndex_(headers);
      values.forEach((row,n) => {
        const firstName = value_(row, idx.firstName);
        const paternal = value_(row, idx.paternal);
        if (!firstName || !paternal) return;
        const maternal = value_(row, idx.maternal);
        const code = value_(row, idx.code);
        output.push({
          id: code || [course,sheetName,n+3].join('|'),
          course: course,
          catechist: sheetName.trim(),
          code: code,
          firstName: firstName,
          paternalLastName: paternal,
          maternalLastName: maternal,
          lastName: [paternal,maternal].filter(Boolean).join(' '),
          birthDate: dateText_(idx.birthDate >= 0 ? row[idx.birthDate] : ''),
          tutor: value_(row, idx.tutor),
          phone: value_(row, idx.phone),
          payment: number_(idx.payment >= 0 ? row[idx.payment] : 0),
          notes: value_(row, idx.notes),
          paymentStatus: value_(row, idx.paymentStatus),
          sourceRow: n + 3,
          updatedAt: new Date().toISOString()
        });
      });
    });
  });
  return output;
}

function upsertStudent_(student) {
  const course = canonicalCourse_(student.course);
  const book = BOOKS[course];
  if (!book) throw new Error('Curso no autorizado');

  const requested = String(student.catechist || '').trim();
  const sheetName = book.sheets.find(x => normalize_(x) === normalize_(requested));
  if (!sheetName) throw new Error('Catequista no autorizado para ' + course);

  const ss = SpreadsheetApp.openById(book.id);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('No se encontró la pestaña ' + sheetName);

  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(2,1,1,lastCol).getDisplayValues()[0];
  const idx = headerIndex_(headers);
  const code = String(student.code || '').trim();
  let rowNumber = findByCode_(sh, idx.code, code);

  if (!rowNumber) {
    rowNumber = firstBlankRow_(sh, idx.paternal);
    prepareRow_(sh, rowNumber, lastCol);
  }

  const split = splitLastName_(student);
  setInput_(sh,rowNumber,idx.paternal,split.paternal);
  setInput_(sh,rowNumber,idx.maternal,split.maternal);
  setInput_(sh,rowNumber,idx.firstName,student.firstName);
  setInput_(sh,rowNumber,idx.birthDate,parseDate_(student.birthDate));
  setInput_(sh,rowNumber,idx.tutor,student.tutor);
  setInput_(sh,rowNumber,idx.phone,String(student.phone || ''));
  setInput_(sh,rowNumber,idx.payment,number_(student.payment));
  setInput_(sh,rowNumber,idx.notes,student.notes);
  if (code && idx.code >= 0 && !sh.getRange(rowNumber,idx.code+1).getFormula()) {
    setInput_(sh,rowNumber,idx.code,code);
  }
  SpreadsheetApp.flush();

  return {
    id: code || [course,sheetName,rowNumber].join('|'),
    course: course,
    catechist: sheetName.trim(),
    code: idx.code >= 0 ? sh.getRange(rowNumber,idx.code+1).getDisplayValue() : code,
    firstName: String(student.firstName || '').trim(),
    paternalLastName: split.paternal,
    maternalLastName: split.maternal,
    lastName: [split.paternal,split.maternal].filter(Boolean).join(' '),
    birthDate: dateText_(parseDate_(student.birthDate)),
    tutor: String(student.tutor || '').trim(),
    phone: String(student.phone || '').trim(),
    payment: number_(student.payment),
    notes: String(student.notes || '').trim(),
    sourceRow: rowNumber,
    updatedAt: new Date().toISOString()
  };
}

function headerIndex_(headers) {
  const h = headers.map(normalize_);
  const find = names => h.findIndex(x => names.some(n => x.indexOf(n) >= 0));
  return {
    paternal: find(['APELLIDO PATERNO']),
    maternal: find(['APELLIDO MATERNO']),
    firstName: find(['NOMBRE (S)','NOMBRES']),
    birthDate: find(['FECHA DE NAC']),
    tutor: find(['TUTOR']),
    phone: find(['WHATSAPP','TELEFONO']),
    payment: find(['PAGO']),
    notes: find(['OBSERVACIONES']),
    paymentStatus: find(['STATUS DE PAGO','ESTADO DE PAGO']),
    code: find(['CODIGO QR FIJO','CODIGO QR','ID'])
  };
}

function findByCode_(sh, codeIndex, code) {
  if (codeIndex < 0 || !code) return 0;
  const count = Math.max(0, sh.getLastRow()-2);
  if (!count) return 0;
  const values = sh.getRange(3,codeIndex+1,count,1).getDisplayValues();
  const target = normalizeCode_(code);
  for (let i=0;i<values.length;i++) {
    if (normalizeCode_(values[i][0]) === target) return i+3;
  }
  return 0;
}

function firstBlankRow_(sh, paternalIndex) {
  if (paternalIndex < 0) throw new Error('La pestaña no tiene APELLIDO PATERNO');
  const max = Math.max(sh.getMaxRows(),3);
  const values = sh.getRange(3,paternalIndex+1,max-2,1).getDisplayValues();
  for (let i=0;i<values.length;i++) if (!String(values[i][0]).trim()) return i+3;
  sh.insertRowAfter(max);
  return max+1;
}

function prepareRow_(sh,row,lastCol) {
  if (row <= 3) return;
  const source = sh.getRange(row-1,1,1,lastCol);
  const target = sh.getRange(row,1,1,lastCol);
  source.copyTo(target,SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);
  source.copyTo(target,SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,false);
  const formulas = source.getFormulas()[0];
  formulas.forEach((formula,i) => {
    if (formula) source.getCell(1,i+1).copyTo(target.getCell(1,i+1),SpreadsheetApp.CopyPasteType.PASTE_FORMULA,false);
  });
  if (!target.getCell(1,1).getFormula()) target.getCell(1,1).setValue(row-2);
}

function setInput_(sh,row,index,value) {
  if (index < 0 || value === undefined) return;
  const cell = sh.getRange(row,index+1);
  if (cell.getFormula()) return;
  cell.setValue(value === null ? '' : value);
}

function splitLastName_(student) {
  if (student.paternalLastName || student.maternalLastName) {
    return {paternal:String(student.paternalLastName||'').trim(),maternal:String(student.maternalLastName||'').trim()};
  }
  const parts = String(student.lastName || '').trim().split(/\s+/).filter(Boolean);
  return {paternal:parts.shift() || '',maternal:parts.join(' ')};
}

function canonicalCourse_(course) {
  const n = normalize_(course);
  return Object.keys(BOOKS).find(x => normalize_(x) === n) || '';
}

function publicConfig_() {
  const out = {};
  Object.keys(BOOKS).forEach(k => out[k] = BOOKS[k].sheets.map(x => x.trim()));
  return out;
}

function log_(action,course,catechist,code,result) {
  try {
    const ss = SpreadsheetApp.openById(BOOKS['INICIACIÓN'].id);
    let sh = ss.getSheetByName('_SYNC_LOG');
    if (!sh) {
      sh = ss.insertSheet('_SYNC_LOG');
      sh.hideSheet();
      sh.appendRow(['FECHA','ACCION','CURSO','CATEQUISTA','CODIGO','RESULTADO']);
    }
    sh.appendRow([new Date(),action,course,catechist,code,result]);
  } catch (_) {}
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function normalize_(value) {
  return String(value == null ? '' : value).trim().toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'');
}
function normalizeCode_(value) { return normalize_(value).replace(/\s+/g,''); }
function value_(row,index) { return index < 0 ? '' : String(row[index] == null ? '' : row[index]).trim(); }
function number_(value) { const n = Number(value); return isFinite(n) ? n : 0; }
function dateText_(value) {
  if (!value) return '';
  const d = value instanceof Date ? value : new Date(value);
  return isNaN(d.getTime()) ? String(value) : Utilities.formatDate(d,'America/Mexico_City','yyyy-MM-dd');
}
function parseDate_(value) {
  if (!value) return '';
  const parts = String(value).slice(0,10).split('-').map(Number);
  return parts.length === 3 && parts.every(Boolean) ? new Date(parts[0],parts[1]-1,parts[2]) : value;
}
