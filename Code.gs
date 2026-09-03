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
 *  POST {action:'reopen_wizyta', id}   (cofa zakończenie protokołu z powrotem do w_toku)
 *  POST {action:'delete_wizyta', id}
 *  POST {action:'analyze_photo', imageBase64, mimeType}
 *  POST {action:'add_urzadzenie', wizytaId, producent, model, sn, typ, system, lokalizacja, uwagi,
 *        czynnik, czynnikFabryczny, czynnikDodatkowy (kg czynnika chłodniczego — tylko jednostki zewnętrzne),
 *        imageBase64, mimeType}
 *  POST {action:'update_urzadzenie', id, ...pola do zmiany (w tym parujZId — ID jednostki
 *        zewnętrznej, do której przypięta jest wewnętrzna), imageBase64 (opcjonalnie nowe zdjęcie)}
 *  POST {action:'delete_urzadzenie', id}
 *  POST {action:'reorder_urzadzenia', wizytaId, orderedIds:[...]}
 *  POST {action:'add_zdjecie', urzadzenieId, imageBase64, mimeType}   (zdjęcie dodatkowe, bez odczytu AI)
 *  POST {action:'delete_zdjecie', id}
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
    s.appendRow(['ID', 'WizytaID', 'Kolejnosc', 'Producent', 'Model', 'SN', 'Typ', 'Lokalizacja', 'Uwagi', 'ZdjecieUrl', 'ZdjecieFileId', 'DataUtworzenia', 'System', 'ParujZId', 'Czynnik', 'CzynnikFabryczny', 'CzynnikDodatkowy']);
  } else {
    // Migracja arkusza założonego przed dodaniem kolejnych kolumn.
    if (s.getLastColumn() < 13) s.getRange(1, 13).setValue('System');
    if (s.getLastColumn() < 14) s.getRange(1, 14).setValue('ParujZId');
    if (s.getLastColumn() < 15) s.getRange(1, 15).setValue('Czynnik');
    if (s.getLastColumn() < 16) s.getRange(1, 16).setValue('CzynnikFabryczny');
    if (s.getLastColumn() < 17) s.getRange(1, 17).setValue('CzynnikDodatkowy');
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

// Cofnięcie zakończenia — na pomyłki przy klikaniu, albo gdy trzeba coś
// jeszcze doprawić już po zamknięciu protokołu. Czyści datę zakończenia,
// żeby nie zostawała nieaktualna do czasu ponownego "Zakończ".
function reopenWizyta(id) {
  const sheet = getWizytySheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.getRange(i + 1, 4).setValue('w_toku');
      sheet.getRange(i + 1, 7).setValue('');
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

// ============ ZDJĘCIA DODATKOWE (filtry, pompka skroplin, sprężarka itd.) ============
function getZdjeciaSheet() {
  const s = getSheet('Zdjecia');
  if (s.getLastRow() === 0) s.appendRow(['ID', 'UrzadzenieId', 'Url', 'FileId', 'DataUtworzenia']);
  return s;
}
function rowToZdjecie(r) {
  return { id: r[0], urzadzenieId: r[1], url: r[2], fileId: r[3], dataUtworzenia: r[4] };
}
function listZdjeciaByUrzadzenie() {
  const rows = getZdjeciaSheet().getDataRange().getValues();
  const byUrz = {};
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    const uid = rows[i][1];
    if (!byUrz[uid]) byUrz[uid] = [];
    byUrz[uid].push(rowToZdjecie(rows[i]));
  }
  return byUrz;
}
// Ten sam folder co tabliczka (Klient/Obiekt-data) — celowo, żeby wszystkie
// zdjęcia jednego urządzenia leżały razem. Nazwa pliku dostaje sufiks z
// godziną, bo w odróżnieniu od tabliczki może być ich wiele na urządzenie.
function saveZdjecieDodatkowe(wizyta, dane, base64, mimeType) {
  const root = getOrCreateFolder(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);
  const klientFolder = getOrCreateFolder(root, wizyta.klient);
  const dataKrotka = Utilities.formatDate(new Date(wizyta.dataUtworzenia), 'Europe/Warsaw', 'yyyy-MM-dd');
  const obiektFolder = getOrCreateFolder(klientFolder, wizyta.obiekt + ' — ' + dataKrotka);

  const ext = (mimeType && mimeType.indexOf('png') >= 0) ? 'png' : 'jpg';
  const opis = [dane.lokalizacja, dane.producent, dane.model].filter(x => x).join(' - ');
  const stamp = Utilities.formatDate(new Date(), 'Europe/Warsaw', 'HHmmss');
  const fileName = sanitizeName(opis || 'urzadzenie') + ' - dodatkowe ' + stamp + '.' + ext;

  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType || 'image/jpeg', fileName);
  const file = obiektFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { url: file.getUrl(), fileId: file.getId() };
}
function addZdjecie(data) {
  if (!data.imageBase64) throw new Error('Brak zdjęcia');
  const uSheet = getUrzadzeniaSheet();
  const uRows = uSheet.getDataRange().getValues();
  let urzRow = null;
  for (let i = 1; i < uRows.length; i++) {
    if (String(uRows[i][0]) === String(data.urzadzenieId)) { urzRow = uRows[i]; break; }
  }
  if (!urzRow) throw new Error('Nie znaleziono urządzenia');
  const wizyta = getWizyta(urzRow[1]);
  if (!wizyta) throw new Error('Nie znaleziono wizyty');
  if (wizyta.status !== 'w_toku') throw new Error('Ta wizyta jest już zakończona');

  const dane = { lokalizacja: urzRow[7], producent: urzRow[3], model: urzRow[4] };
  const foto = saveZdjecieDodatkowe(wizyta, dane, data.imageBase64, data.mimeType);
  const id = newId('z');
  const ts = nowStr();
  getZdjeciaSheet().appendRow([id, data.urzadzenieId, foto.url, foto.fileId, ts]);
  touchWizyta(urzRow[1]);
  return { id: id, urzadzenieId: data.urzadzenieId, url: foto.url, fileId: foto.fileId, dataUtworzenia: ts };
}
function deleteZdjecie(id) {
  const sheet = getZdjeciaSheet();
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

// ============ URZĄDZENIA ============
function rowToUrzadzenie(r) {
  return {
    id: r[0], wizytaId: r[1], kolejnosc: r[2], producent: r[3], model: r[4], sn: r[5],
    typ: r[6], lokalizacja: r[7], uwagi: r[8], zdjecieUrl: r[9], zdjecieFileId: r[10], dataUtworzenia: r[11],
    system: r[12] || '', parujZId: r[13] || '',
    czynnik: r[14] || '', czynnikFabryczny: r[15] || '', czynnikDodatkowy: r[16] || ''
  };
}

function listUrzadzenia(wizytaId) {
  const rows = getUrzadzeniaSheet().getDataRange().getValues();
  const zdjeciaByUrz = listZdjeciaByUrzadzenie();
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    if (!rows[i][0]) continue;
    if (String(rows[i][1]) !== String(wizytaId)) continue;
    const u = rowToUrzadzenie(rows[i]);
    u.zdjeciaDodatkowe = zdjeciaByUrz[u.id] || [];
    out.push(u);
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
    data.typ || '', data.lokalizacja || '', data.uwagi || '', foto.url, foto.fileId, ts, data.system || '', data.parujZId || '',
    data.czynnik || '', data.czynnikFabryczny || '', data.czynnikDodatkowy || ''
  ]);
  touchWizyta(data.wizytaId);
  return {
    id: id, wizytaId: data.wizytaId, kolejnosc: kolejnosc, producent: data.producent || '', model: data.model || '',
    sn: data.sn || '', typ: data.typ || '', lokalizacja: data.lokalizacja || '', uwagi: data.uwagi || '',
    zdjecieUrl: foto.url, zdjecieFileId: foto.fileId, dataUtworzenia: ts, system: data.system || '', parujZId: data.parujZId || '',
    czynnik: data.czynnik || '', czynnikFabryczny: data.czynnikFabryczny || '', czynnikDodatkowy: data.czynnikDodatkowy || ''
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
    if (data.czynnik !== undefined) sheet.getRange(rowNum, 15).setValue(data.czynnik);
    if (data.czynnikFabryczny !== undefined) sheet.getRange(rowNum, 16).setValue(data.czynnikFabryczny);
    if (data.czynnikDodatkowy !== undefined) sheet.getRange(rowNum, 17).setValue(data.czynnikDodatkowy);

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
    return rowToUrzadzenie(sheet.getRange(rowNum, 1, 1, 17).getValues()[0]);
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
    '(Manufacturer/Brand), model (Model/Model No./Type), numer seryjny (Serial No./S/N) oraz — jeśli widoczny — ' +
    'czynnik chłodniczy (Refrigerant/Refrigerant Type/Czynnik chłodniczy, np. R32, R410A, R404A, R134A). ' +
    'WAŻNE — numer seryjny: na wielu tabliczkach numer seryjny nie ma żadnej etykiety "S/N" ani "Serial No." ' +
    'obok siebie — to po prostu ciąg cyfr/znaków wydrukowany bezpośrednio POD kodem kreskowym (lub kodem QR), ' +
    'bez żadnego opisu. Jeśli nie widzisz pola opisanego wprost jako numer seryjny, ale na tabliczce jest kod ' +
    'kreskowy, odczytaj jako numer seryjny ciąg znaków wydrukowany bezpośrednio pod (lub bezpośrednio obok) tego ' +
    'kodu kreskowego — to niemal zawsze jest właśnie numer seryjny, nawet bez etykiety. Jeśli na tabliczce jest ' +
    'kilka kodów kreskowych z ciągami znaków, wybierz ten, który wygląda na numer seryjny (zwykle unikalny, ' +
    'dłuższy ciąg cyfr/liter — nie kod modelu, który już odczytałeś osobno w polu model). WAŻNE — czynnik ' +
    'chłodniczy: pole "czynnik" wypełniaj tylko wtedy, gdy na tabliczce faktycznie widać jego oznaczenie (typowe ' +
    'dla jednostek zewnętrznych/agregatów; jednostki wewnętrzne zwykle go nie pokazują — wtedy zostaw puste). ' +
    'Podaj tam wyłącznie sam symbol czynnika (np. "R32"), bez GWP ani ilości w kg — te ewentualnie widoczne obok ' +
    'liczby (np. "0,62 kg") pomiń. Zwróć WYŁĄCZNIE czysty JSON, bez dodatkowego tekstu, w formacie ' +
    '{"producent":"...","model":"...","sn":"...","czynnik":"..."}. Jeśli któregoś pola nie da się odczytać, ' +
    'wstaw pusty string "". Przepisz model, numer seryjny i czynnik dokładnie tak, znak po znaku, jak są na ' +
    'tabliczce — nie poprawiaj ich i nie zgaduj.';

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + GEMINI_MODEL + ':generateContent?key=' + encodeURIComponent(key);
  const payload = {
    contents: [{ parts: [{ text: prompt }, { inline_data: { mime_type: mimeType || 'image/jpeg', data: base64 } }] }],
    generationConfig: { temperature: 0, responseMimeType: 'application/json' }
  };

  // 503 (model chwilowo przeciążony) i 429 (limit zapytań) zwykle mijają po
  // chwili — próbujemy ponownie zamiast od razu poddawać serwisanta ręcznemu
  // wpisywaniu. Inne kody (zły klucz, zła nazwa modelu) ponowna próba nie naprawi.
  const MAX_TRIES = 3;
  let code, bodyText;
  for (let attempt = 1; attempt <= MAX_TRIES; attempt++) {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    code = res.getResponseCode();
    bodyText = res.getContentText();
    if (code >= 200 && code < 300) break;
    if ((code === 503 || code === 429) && attempt < MAX_TRIES) {
      Utilities.sleep(1200 * attempt);
      continue;
    }
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
  try { out = JSON.parse(text); } catch (e) { out = { producent: '', model: '', sn: '', czynnik: '' }; }
  return { producent: out.producent || '', model: out.model || '', sn: out.sn || '', czynnik: out.czynnik || '' };
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
    if (action === 'reopen_wizyta') return jsonOut({ ok: true, done: reopenWizyta(body.id) });
    if (action === 'delete_wizyta') return jsonOut({ ok: true, deleted: deleteWizyta(body.id) });
    if (action === 'analyze_photo') return jsonOut({ ok: true, data: analyzePhoto(body.imageBase64, body.mimeType) });
    if (action === 'add_urzadzenie') return jsonOut({ ok: true, data: addUrzadzenie(body) });
    if (action === 'update_urzadzenie') return jsonOut({ ok: true, data: updateUrzadzenie(body) });
    if (action === 'delete_urzadzenie') return jsonOut({ ok: true, deleted: deleteUrzadzenie(body.id) });
    if (action === 'add_zdjecie') return jsonOut({ ok: true, data: addZdjecie(body) });
    if (action === 'delete_zdjecie') return jsonOut({ ok: true, deleted: deleteZdjecie(body.id) });
    if (action === 'reorder_urzadzenia') return jsonOut({ ok: true, done: reorderUrzadzenia(body.wizytaId, body.orderedIds || []) });
    return jsonOut({ ok: false, error: 'Nieznana akcja' });
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}
