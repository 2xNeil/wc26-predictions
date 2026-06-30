const SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet(e) {
  const action = e.parameter.action;
  const ss = SpreadsheetApp.openById(SS_ID);

  if (action === 'games') {
    return json(sheetToJson(ss.getSheetByName('games')));
  }

  if (action === 'predictions') {
    return json(sheetToJson(ss.getSheetByName('predictions')));
  }

  if (action === 'verifyPin') {
    const { user, pin } = e.parameter;
    const rows = ss.getSheetByName('users').getDataRange().getValues().slice(1);
    const match = rows.find(r => r[0] === user && String(r[1]) === String(pin));
    return json({ valid: !!match });
  }

  if (action === 'predict') {
    const { user, pin, gameId, winner, score1, score2 } = e.parameter;

    // Verify PIN
    const userRows = ss.getSheetByName('users').getDataRange().getValues().slice(1);
    const validUser = userRows.find(r => r[0] === user && String(r[1]) === String(pin));
    if (!validUser) return json({ error: 'Invalid PIN' });

    // One-time lock: don't overwrite existing prediction
    const predSheet = ss.getSheetByName('predictions');
    const predRows = predSheet.getDataRange().getValues().slice(1);
    const exists = predRows.find(r => String(r[0]) === String(gameId) && r[1] === user);
    if (exists) return json({ error: 'Already locked in' });

    predSheet.appendRow([gameId, user, winner, Number(score1), Number(score2), new Date().toISOString()]);
    return json({ success: true });
  }

  if (action === 'submitPicks') {
    const { user, pin, picks } = e.parameter;

    // Verify PIN once
    const userRows = ss.getSheetByName('users').getDataRange().getValues().slice(1);
    const validUser = userRows.find(r => r[0] === user && String(r[1]) === String(pin));
    if (!validUser) return json({ error: 'Invalid PIN' });

    let picksArr;
    try { picksArr = JSON.parse(picks); }
    catch (_) { return json({ error: 'Invalid picks format' }); }

    // Read existing predictions once
    const predSheet = ss.getSheetByName('predictions');
    const predRows = predSheet.getDataRange().getValues().slice(1);
    const existing = new Set(predRows.map(r => `${r[0]}_${r[1]}`));

    const now = new Date().toISOString();
    const saved = [], skipped = [], rowsToWrite = [];

    for (const pick of picksArr) {
      const key = `${pick.gameId}_${user}`;
      if (existing.has(key)) { skipped.push(String(pick.gameId)); continue; }
      rowsToWrite.push([pick.gameId, user, pick.winner, Number(pick.score1), Number(pick.score2), now]);
      existing.add(key);
      saved.push(String(pick.gameId));
    }

    // Single bulk write instead of N appendRow calls
    if (rowsToWrite.length) {
      predSheet.getRange(predSheet.getLastRow() + 1, 1, rowsToWrite.length, rowsToWrite[0].length)
              .setValues(rowsToWrite);
    }

    return json({ success: true, saved, skipped });
  }

  return json({ error: 'Unknown action' });
}

function sheetToJson(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[String(h)] = row[i]);
    return obj;
  });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}