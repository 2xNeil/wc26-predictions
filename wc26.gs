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

  if (action === 'players') {
    const rows = ss.getSheetByName('users').getDataRange().getValues().slice(1);
    const joined = {};
    for (const r of rows) {
      const u = r[0];
      if (!u) continue;
      const sf = Number(r[3]) || 1;
      joined[u] = joined[u] ? Math.min(joined[u], sf) : sf;
    }
    return json(Object.keys(joined).map(user => ({ user, start_from: joined[user] })));
  }

  if (action === 'verifyPin') {
    const { pin } = e.parameter;
    const rows = ss.getSheetByName('users').getDataRange().getValues().slice(1);
    const match = rows.find(r => String(r[1]) === String(pin));
    if (!match) return json({ valid: false });
    const opponentRow = rows.find(r => r[0] === match[2] && r[2] === match[0]);
    const cutoff = r => Number(r && r[3]) || 1;
    const startFrom = Math.max(cutoff(match), cutoff(opponentRow));
    return json({ valid: true, user: match[0], opponent: match[2], startFrom });
  }

  if (action === 'submitPicks') {
    const { user, pin, picks } = e.parameter;

    // Verify PIN and that it belongs to this user
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
