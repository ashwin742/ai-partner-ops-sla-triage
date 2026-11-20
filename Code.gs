const GEMINI_API_KEY = 'YOUR_GEMINI_API_KEY_HERE';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AI Ops Triage')
    .addItem('Run AI Ticket Triage', 'triagePartnerTickets')
    .addToUi();
}

function triagePartnerTickets() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Ticket_Queue');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Sheet "Ticket_Queue" not found!');
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const data = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.5-flash-lite:generateContent';

  // Collect only untriaged rows
  const pendingRows = [];
  for (let i = 0; i < data.length; i++) {
    const [ticketId, partner, issue, slaHours, , , , , status] = data[i];
    if (ticketId && issue && status !== 'Triaged') {
      pendingRows.push({ rowNum: i + 2, ticketId, partner, issue, slaHours });
    }
  }

  if (pendingRows.length === 0) {
    SpreadsheetApp.getUi().alert('All tickets are already triaged!');
    return;
  }

  // Process in parallel batches of 5 for speed & rate-limit safety
  const BATCH_SIZE = 5;
  for (let b = 0; b < pendingRows.length; b += BATCH_SIZE) {
    const batch = pendingRows.slice(b, b + BATCH_SIZE);

    const requests = batch.map(item => {
      const prompt = `You are a Senior Partner Operations Specialist supporting an app marketplace ecosystem.
Analyze this incoming developer ticket and return ONLY valid JSON with no markdown backticks:
Partner: ${item.partner}
Issue: ${item.issue}
Hours to SLA Breach: ${item.slaHours}

Return JSON with exact keys:
{
  "category": "One of: Featuring Nomination | Policy & App Review | Technical Onboarding | Account & Billing",
  "priority": "One of: P1 - Critical | P2 - High | P3 - Normal",
  "draft_response": "A concise, polite, 2-sentence partner support reply acknowledging the issue and next action."
}`;

      return {
        url: url,
        method: 'post',
        contentType: 'application/json',
        headers: { 'x-goog-api-key': GEMINI_API_KEY },
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { responseMimeType: "application/json" }
        }),
        muteHttpExceptions: true
      };
    });

    const responses = UrlFetchApp.fetchAll(requests);

    responses.forEach((response, idx) => {
      const item = batch[idx];
      const hours = parseFloat(item.slaHours);
      let slaFlag = 'ON TRACK';
      if (!isNaN(hours) && hours <= 4) {
        slaFlag = 'URGENT: SLA BREACH RISK';
      } else if (!isNaN(hours) && hours <= 8) {
        slaFlag = 'WATCHLIST (< 8H)';
      }

      try {
        const rawText = response.getContentText();
        const json = JSON.parse(rawText);

        if (!json.candidates || !json.candidates[0]) {
          const apiErr = json.error ? json.error.message : rawText;
          sheet.getRange(item.rowNum, 9).setValue('API Error: ' + apiErr);
          return;
        }

        const aiText = json.candidates[0].content.parts[0].text;
        const result = JSON.parse(aiText);

        sheet.getRange(item.rowNum, 5).setValue(result.category);
        sheet.getRange(item.rowNum, 6).setValue(result.priority);
        sheet.getRange(item.rowNum, 7).setValue(slaFlag);
        sheet.getRange(item.rowNum, 8).setValue(result.draft_response);
        sheet.getRange(item.rowNum, 9).setValue('Triaged');
        sheet.getRange(item.rowNum, 10).setValue(new Date().toLocaleString());

        if (slaFlag.includes('URGENT')) {
          sheet.getRange(item.rowNum, 1, 1, 10).setBackground('#fce8e6');
        }
      } catch (e) {
        sheet.getRange(item.rowNum, 9).setValue('Error: ' + e.message);
      }
    });

    SpreadsheetApp.flush(); // Updates the sheet visibly after every 5 rows!
  }
}
