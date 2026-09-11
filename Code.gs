
 *  POST {action:'analyze_photo', imageBase64, mimeType}
 *  POST {action:'add_urzadzenie', wizytaId, producent, model, sn, typ, system, lokalizacja, uwagi,
 *        czynnik, czynnikFabryczny, czynnikDodatkowy (kg czynnika chłodniczego — tylko jednostki zewnętrzne),
 *        imageBase64, mimeType}
 *        imageBase64, mimeType, clientId (opcjonalnie — ID nadane w przeglądarce; powtórne wywołanie
 *        z tym samym clientId+wizytaId zwraca już istniejący wpis zamiast tworzyć duplikat)}
 *  POST {action:'update_urzadzenie', id, ...pola do zmiany (w tym parujZId — ID jednostki
 *        zewnętrznej, do której przypięta jest wewnętrzna), imageBase64 (opcjonalnie nowe zdjęcie)}
 *  POST {action:'delete_urzadzenie', id}
function getUrzadzeniaSheet() {
  const s = getSheet('Urzadzenia');
  if (s.getLastRow() === 0) {
    s.appendRow(['ID', 'WizytaID', 'Kolejnosc', 'Producent', 'Model', 'SN', 'Typ', 'Lokalizacja', 'Uwagi', 'ZdjecieUrl', 'ZdjecieFileId', 'DataUtworzenia', 'System', 'ParujZId', 'Czynnik', 'CzynnikFabryczny', 'CzynnikDodatkowy']);
    s.appendRow(['ID', 'WizytaID', 'Kolejnosc', 'Producent', 'Model', 'SN', 'Typ', 'Lokalizacja', 'Uwagi', 'ZdjecieUrl', 'ZdjecieFileId', 'DataUtworzenia', 'System', 'ParujZId', 'Czynnik', 'CzynnikFabryczny', 'CzynnikDodatkowy', 'ClientId']);
  } else {
    // Migracja arkusza założonego przed dodaniem kolejnych kolumn.
    if (s.getLastColumn() < 13) s.getRange(1, 13).setValue('System');
    if (s.getLastColumn() < 15) s.getRange(1, 15).setValue('Czynnik');
    if (s.getLastColumn() < 16) s.getRange(1, 16).setValue('CzynnikFabryczny');
    if (s.getLastColumn() < 17) s.getRange(1, 17).setValue('CzynnikDodatkowy');
    if (s.getLastColumn() < 18) s.getRange(1, 18).setValue('ClientId');
  }
  return s;
}
    id: r[0], wizytaId: r[1], kolejnosc: r[2], producent: r[3], model: r[4], sn: r[5],
    typ: r[6], lokalizacja: r[7], uwagi: r[8], zdjecieUrl: r[9], zdjecieFileId: r[10], dataUtworzenia: r[11],
    system: r[12] || '', parujZId: r[13] || '',
    czynnik: r[14] || '', czynnikFabryczny: r[15] || '', czynnikDodatkowy: r[16] || ''
    czynnik: r[14] || '', czynnikFabryczny: r[15] || '', czynnikDodatkowy: r[16] || '',
    clientId: r[17] || ''
  };
}

// Szuka urządzenia z tej samej wizyty i tym samym clientId (ID nadanym przez
// przeglądarkę przy dodaniu, zanim serwer zdążył przydzielić własne ID). Służy
// wyłącznie do ochrony przed duplikatami — patrz addUrzadzenie().
function findUrzadzenieByClientId(wizytaId, clientId) {
  const rows = getUrzadzeniaSheet().getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]) === String(wizytaId) && rows[i][17] && String(rows[i][17]) === String(clientId)) {
      return rowToUrzadzenie(rows[i]);
    }
  }
  return null;
}

function listUrzadzenia(wizytaId) {
  const rows = getUrzadzeniaSheet().getDataRange().getValues();
  const zdjeciaByUrz = listZdjeciaByUrzadzenie();
  if (!wizyta) throw new Error('Nie znaleziono wizyty');
  if (wizyta.status !== 'w_toku') throw new Error('Ta wizyta jest już zakończona');

  // Ochrona przed duplikatami: appka w terenie potrafi wysłać to samo nowe
  // urządzenie dwa razy przy słabym zasięgu (np. automatyczna synchronizacja
  // przy otwarciu wizyty nakłada się na tę, która wystrzeliła zaraz po
  // dodaniu). clientId to ID nadane po stronie przeglądarki w momencie
  // dodania — blokada + sprawdzenie po nim gwarantuje jeden wiersz, nawet
  // gdy oba żądania trafią na serwer prawie równocześnie.
  let lock = null;
  if (data.clientId) {
    lock = LockService.getScriptLock();
    lock.waitLock(10000);
    const existing = findUrzadzenieByClientId(data.wizytaId, data.clientId);
    if (existing) { lock.releaseLock(); return existing; }
  }
  try {
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
    const id = newId('u');
    const kolejnosc = data.kolejnosc || nextKolejnosc(data.wizytaId);
    const foto = saveZdjecie(wizyta, data, data.imageBase64, data.mimeType);
    const ts = nowStr();
    getUrzadzeniaSheet().appendRow([
      id, data.wizytaId, kolejnosc, data.producent || '', data.model || '', data.sn || '',
      data.typ || '', data.lokalizacja || '', data.uwagi || '', foto.url, foto.fileId, ts, data.system || '', data.parujZId || '',
      data.czynnik || '', data.czynnikFabryczny || '', data.czynnikDodatkowy || '', data.clientId || ''
    ]);
    touchWizyta(data.wizytaId);
    return {
      id: id, wizytaId: data.wizytaId, kolejnosc: kolejnosc, producent: data.producent || '', model: data.model || '',
      sn: data.sn || '', typ: data.typ || '', lokalizacja: data.lokalizacja || '', uwagi: data.uwagi || '',
      zdjecieUrl: foto.url, zdjecieFileId: foto.fileId, dataUtworzenia: ts, system: data.system || '', parujZId: data.parujZId || '',
      czynnik: data.czynnik || '', czynnikFabryczny: data.czynnikFabryczny || '', czynnikDodatkowy: data.czynnikDodatkowy || '',
      clientId: data.clientId || ''
    };
  } finally {
    if (lock) lock.releaseLock();
  }
}

function updateUrzadzenie(data) {
    }

    touchWizyta(wizytaId);
    return rowToUrzadzenie(sheet.getRange(rowNum, 1, 1, 17).getValues()[0]);
    return rowToUrzadzenie(sheet.getRange(rowNum, 1, 1, 18).getValues()[0]);
  }
  throw new Error('Nie znaleziono urządzenia');
}
