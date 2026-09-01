/**
 * HVAC Notatki — backend
 *
 * Wklej ten kod w: Arkusz Google (nowy, pusty) -> Rozszerzenia -> Apps Script -> wklej -> zapisz.
 * Przed wdrożeniem: Ustawienia projektu (ikona koła zębatego) -> Właściwości skryptu -> Dodaj właściwość:
 *   GEMINI_API_KEY = twój klucz z https://aistudio.google.com/apikey
 * Wdroż jako aplikację internetową: Wykonaj jako "Ja", Kto ma dostęp "Każdy".
 *
 * Endpointy:
 *  GET  ?action=list_wizyty&status=w_toku|zakonczona   (status opcjonalny — brak = wszystkie)
 *  GET  ?action=get_wizyta&id=XXX                       (wizyta + jej urządzenia)
 *  POST {action:'create_wizyta', klient, obiekt}
 *  POST {action:'finish_wizyta', id}
 *  POST {action:'delete_wizyta', id}
 *  POST {action:'analyze_photo', imageBase64, mimeType}
 *  POST {action:'add_urzadzenie', wizytaId, producent, model, sn, typ, system, lokalizacja, uwagi, imageBase64, mimeType}
 *  POST {action:'update_urzadzenie', id, ...pola do zmiany (w tym parujZId — ID jednostki
 *        zewnętrznej, do której przypięta jest wewnętrzna), imageBase64 (opcjonalnie nowe zdjęcie)}
 *  POST {action:'delete_urzadzenie', id}
 *  POST {action:'reorder_urzadzenia', wizytaId, orderedIds:[...]}
 */

// Musi być identyczne z TOKEN w index.html — bariera przed wywołaniem /exec
// przez kogoś, kto zna sam adres URL, ale nie appkę.
const SHARED_SECRET = 'hvac_580gMoLePXK8REU7k8KfTyG5';

const ROOT_FOLDER_NAME = 'HVAC Notatki';
// Nazwa modelu Gemini — jeśli Google zmieni nazewnictwo i odczyt zacznie zwracać
// błąd 404, podmień tę jedną stałą na aktualną nazwę z https://ai.google.dev/gemini-api/docs/models
// (2.5-flash wycofane dla nowych użytkowników 31.08.2026 — Google w treści błędu 404
// wskazał 3.6-flash jako następcę, stąd ta wartość).
const GEMINI_MODEL = 'gemini-3.6-flash';

function checkAuth(token) {
  if (token !== SHARED_SECRET) throw new Error('Brak autoryzacji');
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function nowStr() {
  return Utilities.formatDate(new Date(), 'Europe/Warsaw', "yyyy-MM-dd'T'HH:mm:ss");
}

function newId(prefix) {
  return prefix + '_' + new Date().getTime() + '_' + Math.floor(Math.random() * 10000);
}

// ============ ARKUSZE ============
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}
function getWizytySheet() {
  const s = getSheet('Wizyty');
  if (s.getLastRow() === 0) s.appendRow(['ID', 'Klient', 'Obiekt', 'Status', 'DataUtworzenia', 'DataAktywnosci', 'DataZakonczenia']);
  return s;
}
function getUrzadzeniaSheet() {
  const s = getSheet('Urzadzenia');
  if (s.getLastRow() === 0) {
    s.appendRow(['ID', 'WizytaID', 'Kolejnosc', 'Producent', 'Model', 'SN', 'Typ', 'Lokalizacja', 'Uwagi', 'ZdjecieUrl', 'ZdjecieFileId', 'DataUtworzenia', 'System', 'ParujZId']);
  } else {
    // Migracja arkusza założonego przed dodaniem kolumn System / ParujZId.
    if (s.getLastColumn() < 13) s.getRange(1, 13).setValue('System');
    if (s.getLastColumn() < 14) s.getRange(1, 14).setValue('ParujZId');
  }
  return s;
}

// ============ WIZYTY ============
function rowToWizyta(r) {
  return { id: r[0], klient: r[1], obiekt: r[2], status: r[3], dataUtworzenia: r[4], dataAktywnosci: r[5], dataZakonczenia: r[6] };
}

function listWizyty(status) {
  const rows = getWizytySheet().getDataRange().getValues();
  const urzRows = getUrzadzeniaSheet().getDataRange().getValues();
  const counts = {};
  for (let i = 1; i < urzRows.length; i++) {
    const wid = urzRows[i][1];
    if (!wid) continue;
    counts[wid] = (counts[wid] || 0) + 1;
  }
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    if (status && rows[i][3] !== status) continue;
    const w = rowToWizyta(rows[i]);
    w.liczbaUrzadzen = counts[w.id] || 0;
    out.push(w);
  }
  out.sort((a, b) => new Date(b.dataAktywnosci) - new Date(a.dataAktywnosci));
  return out;
}

function getWizyta(id) {
  const rows = getWizytySheet().getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) return rowToWizyta(rows[i]);
  }
  return null;
}

function touchWizyta(id) {
  const sheet = getWizytySheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.getRange(i + 1, 6).setValue(nowStr());
      return;
    }
  }
}

function createWizyta(klient, obiekt) {
  klient = String(klient || '').trim();
  obiekt = String(obiekt || '').trim();
  if (!klient || !obiekt) throw new Error('Podaj klienta i obiekt');
  const id = newId('w');
  const ts = nowStr();
  getWizytySheet().appendRow([id, klient, obiekt, 'w_toku', ts, ts, '']);
  return { id: id, klient: klient, obiekt: obiekt, status: 'w_toku', dataUtworzenia: ts, dataAktywnosci: ts, liczbaUrzadzen: 0 };
}

function finishWizyta(id) {
  const sheet = getWizytySheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.getRange(i + 1, 4).setValue('zakonczona');
      sheet.getRange(i + 1, 7).setValue(nowStr());
      return true;
    }
  }
  throw new Error('Nie znaleziono wizyty');
}

// Kasuje wizytę i wszystkie jej urządzenia z arkuszy. Zdjęcia na Dysku zostają
// nietknięte (celowo — usuwanie plików to osobne, bardziej ryzykowne działanie).
function deleteWizyta(id) {
  const uSheet = getUrzadzeniaSheet();
  const uRows = uSheet.getDataRange().getValues();
  for (let i = uRows.length - 1; i >= 1; i--) {
    if (String(uRows[i][1]) === String(id)) uSheet.deleteRow(i + 1);
  }
  const wSheet = getWizytySheet();
  const wRows = wSheet.getDataRange().getValues();
  for (let i = 1; i < wRows.length; i++) {
    if (String(wRows[i][0]) === String(id)) {
      wSheet.deleteRow(i + 1);
      return true;
    }
  }
  throw new Error('Nie znaleziono wizyty');
}

// ============ URZĄDZENIA ============
function rowToUrzadzenie(r) {
  return {
    id: r[0], wizytaId: r[1], kolejnosc: r[2], producent: r[3], model: r[4], sn: r[5],
    typ: r[6], lokalizacja: r[7], uwagi: r[8], zdjecieUrl: r[9], zdjecieFileId: r[10], dataUtworzenia: r[11],
    system: r[12] || '', parujZId: r[13] || ''
  };
}

function listUrzadzenia(wizytaId) {
  const rows = getUrzadzeniaSheet().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    if (String(rows[i][1]) !== String(wizytaId)) continue;
    out.push(rowToUrzadzenie(rows[i]));
  }
  out.sort((a, b) => (Number(a.kolejnosc) || 0) - (Number(b.kolejnosc) || 0));
  return out;
}

function nextKolejnosc(wizytaId) {
  let max = 0;
  listUrzadzenia(wizytaId).forEach(u => { const k = Number(u.kolejnosc) || 0; if (k > max) max = k; });
  return max + 10;
}

function sanitizeName(s) {
  return String(s || '').replace(/[\\\/:*?"<>|]/g, '-').replace(/\s+/g, ' ').trim();
}

function getOrCreateFolder(parent, name) {
  name = sanitizeName(name) || 'Bez nazwy';
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

// Ścieżka: HVAC Notatki / {Klient} / {Obiekt} — {data wizyty} / plik nazwany po lokalizacji + urządzeniu
function saveZdjecie(wizyta, dane, base64, mimeType) {
  if (!base64) return { url: '', fileId: '' };
  const root = getOrCreateFolder(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);
  const klientFolder = getOrCreateFolder(root, wizyta.klient);
  const dataKrotka = Utilities.formatDate(new Date(wizyta.dataUtworzenia), 'Europe/Warsaw', 'yyyy-MM-dd');
  const obiektFolder = getOrCreateFolder(klientFolder, wizyta.obiekt + ' — ' + dataKrotka);

  const ext = (mimeType && mimeType.indexOf('png') >= 0) ? 'png' : 'jpg';
  const opis = [dane.lokalizacja, dane.producent, dane.model, dane.sn].filter(x => x).join(' - ');
  const fileName = sanitizeName(opis || 'urzadzenie') + '.' + ext;

  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', fileName);
  const file = obiektFolder.createFile(blob);
  // Bez tego miniatury w appce byłyby niewidoczne dla serwisantów niezalogowanych
  // na konto Google, na które wdrożony jest ten skrypt.
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { url: file.getUrl(), fileId: file.getId() };
}

function addUrzadzenie(data) {
  const wizyta = getWizyta(data.wizytaId);
  if (!wizyta) throw new Error('Nie znaleziono wizyty');
  if (wizyta.status !== 'w_toku') throw new Error('Ta wizyta jest już zakończona');

  const id = newId('u');
  const kolejnosc = data.kolejnosc || nextKolejnosc(data.wizytaId);
  const foto = saveZdjecie(wizyta, data, data.imageBase64, data.mimeType);
  const ts = nowStr();
  getUrzadzeniaSheet().appendRow([
    id, data.wizytaId, kolejnosc, data.producent || '', data.model || '', data.sn || '',
    data.typ || '', data.lokalizacja || '', data.uwagi || '', foto.url, foto.fileId, ts, data.system || '', data.parujZId || ''
  ]);
  touchWizyta(data.wizytaId);
  return {
    id: id, wizytaId: data.wizytaId, kolejnosc: kolejnosc, producent: data.producent || '', model: data.model || '',
    sn: data.sn || '', typ: data.typ || '', lokalizacja: data.lokalizacja || '', uwagi: data.uwagi || '',
    zdjecieUrl: foto.url, zdjecieFileId: foto.fileId, dataUtworzenia: ts, system: data.system || '', parujZId: data.parujZId || ''
  };
}

function updateUrzadzenie(data) {
  const sheet = getUrzadzeniaSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) !== String(data.id)) continue;

    const wizytaId = rows[i][1];
    const wizyta = getWizyta(wizytaId);
    if (!wizyta) throw new Error('Nie znaleziono wizyty');
    if (wizyta.status !== 'w_toku') throw new Error('Ta wizyta jest już zakończona');

    const rowNum = i + 1;
    if (data.producent !== undefined) sheet.getRange(rowNum, 4).setValue(data.producent);
    if (data.model !== undefined) sheet.getRange(rowNum, 5).setValue(data.model);
    if (data.sn !== undefined) sheet.getRange(rowNum, 6).setValue(data.sn);
    if (data.typ !== undefined) sheet.getRange(rowNum, 7).setValue(data.typ);
    if (data.lokalizacja !== undefined) sheet.getRange(rowNum, 8).setValue(data.lokalizacja);
    if (data.uwagi !== undefined) sheet.getRange(rowNum, 9).setValue(data.uwagi);
    if (data.system !== undefined) sheet.getRange(rowNum, 13).setValue(data.system);
    if (data.parujZId !== undefined) sheet.getRange(rowNum, 14).setValue(data.parujZId);

    if (data.imageBase64) {
      const merged = {
        lokalizacja: data.lokalizacja !== undefined ? data.lokalizacja : rows[i][7],
        producent: data.producent !== undefined ? data.producent : rows[i][3],
        model: data.model !== undefined ? data.model : rows[i][4],
        sn: data.sn !== undefined ? data.sn : rows[i][5]
      };
      const foto = saveZdjecie(wizyta, merged, data.imageBase64, data.mimeType);
      sheet.getRange(rowNum, 10).setValue(foto.url);
      sheet.getRange(rowNum, 11).setValue(foto.fileId);
    }

    touchWizyta(wizytaId);
    return rowToUrzadzenie(sheet.getRange(rowNum, 1, 1, 14).getValues()[0]);
  }
  throw new Error('Nie znaleziono urządzenia');
}

function deleteUrzadzenie(id) {
  const sheet = getUrzadzeniaSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      const wizytaId = rows[i][1];
      const fileId = rows[i][10];
      sheet.deleteRow(i + 1);
      if (fileId) { try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) { } }
      touchWizyta(wizytaId);
      return true;
    }
  }
  return false;
}

function reorderUrzadzenia(wizytaId, orderedIds) {
  const sheet = getUrzadzeniaSheet();
  const rows = sheet.getDataRange().getValues();
  const pos = {};
  (orderedIds || []).forEach((id, idx) => { pos[id] = (idx + 1) * 10; });
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) !== String(wizytaId)) continue;
    const id = rows[i][0];
    if (pos[id] !== undefined) sheet.getRange(i + 1, 3).setValue(pos[id]);
  }
  touchWizyta(wizytaId);
  return true;
}

// ============ ODCZYT TABLICZKI (Gemini) ============
function analyzePhoto(base64, mimeType) {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('Brak klucza GEMINI_API_KEY — dodaj go w Ustawienia projektu -> Właściwości skryptu');

  const prompt = 'Jesteś asystentem serwisanta klimatyzacji/HVAC. Na zdjęciu jest tabliczka znamionowa ' +
    'urządzenia (klimatyzator, centrala wentylacyjna, agregat wody lodowej itp). Odczytaj z niej: producenta ' +
    '(Manufacturer/Brand), model (Model/Model No./Type) i numer seryjny (Serial No./S/N). Zwróć WYŁĄCZNIE ' +
    'czysty JSON, bez dodatkowego tekstu, w formacie {"producent":"...","model":"...","sn":"..."}. Jeśli ' +
    'któregoś pola nie da się odczytać, wstaw pusty string "". Przepisz model i numer seryjny dokładnie tak, ' +
    'znak po znaku, jak są na tabliczce — nie poprawiaj ich i nie zgaduj.';

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + encodeURIComponent(key);
  const payload = {
    contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: mimeType || 'image/jpeg', data: base64 } }] }],
    generationConfig: { temperature: 0, responseMimeType: 'application/json' }
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  const bodyText = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Błąd odczytu AI (' + code + '): ' + bodyText.slice(0, 300));
  }
  const parsed = JSON.parse(bodyText);
  let text;
  try {
    text = parsed.candidates[0].content.parts[0].text;
  } catch (e) {
    throw new Error('Nieoczekiwana odpowiedź AI');
  }
  text = text.trim().replace(/^```json/i, '').replace(/^```/, '').replace(/```$/, '').trim();
  let out;
  try { out = JSON.parse(text); } catch (e) { out = { producent: '', model: '', sn: '' }; }
  return { producent: out.producent || '', model: out.model || '', sn: out.sn || '' };
}

// ============ ROUTER ============
function doGet(e) {
  try {
    checkAuth(e.parameter.token);
    const action = e.parameter.action;
    if (action === 'list_wizyty') {
      return jsonOut({ ok: true, data: listWizyty(e.parameter.status) });
    }
    if (action === 'get_wizyta') {
      const w = getWizyta(e.parameter.id);
      if (!w) return jsonOut({ ok: false, error: 'Nie znaleziono wizyty' });
      w.urzadzenia = listUrzadzenia(e.parameter.id);
      return jsonOut({ ok: true, data: w });
    }
    return jsonOut({ ok: false, error: 'Nieznana akcja' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    checkAuth(body.token);
    const action = body.action;
    if (action === 'create_wizyta') return jsonOut({ ok: true, data: createWizyta(body.klient, body.obiekt) });
    if (action === 'finish_wizyta') return jsonOut({ ok: true, done: finishWizyta(body.id) });
    if (action === 'delete_wizyta') return jsonOut({ ok: true, deleted: deleteWizyta(body.id) });
    if (action === 'analyze_photo') return jsonOut({ ok: true, data: analyzePhoto(body.imageBase64, body.mimeType) });
    if (action === 'add_urzadzenie') return jsonOut({ ok: true, data: addUrzadzenie(body) });
    if (action === 'update_urzadzenie') return jsonOut({ ok: true, data: updateUrzadzenie(body) });
    if (action === 'delete_urzadzenie') return jsonOut({ ok: true, deleted: deleteUrzadzenie(body.id) });
    if (action === 'reorder_urzadzenia') return jsonOut({ ok: true, done: reorderUrzadzenia(body.wizytaId, body.orderedIds || []) });
    return jsonOut({ ok: false, error: 'Nieznana akcja' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}
