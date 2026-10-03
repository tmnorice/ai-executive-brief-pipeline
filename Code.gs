const BRIEF_CONFIG = {
  QUEUE_SHEET: 'Briefing Queue',
  SOURCE_SHEET: 'Source Log',
  TIME_ZONE: 'America/Phoenix'
};

function pullCalendarContext() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = spreadsheet.getSheetByName(BRIEF_CONFIG.QUEUE_SHEET);
  const sourceSheet = spreadsheet.getSheetByName(BRIEF_CONFIG.SOURCE_SHEET);

  if (!queueSheet || !sourceSheet) {
    throw new Error('Required sheet not found.');
  }

  const queueData = queueSheet.getDataRange().getValues();
  const headers = queueData[0];

  function col(name) {
    const index = headers.indexOf(name);

    if (index === -1) {
      throw new Error(`Missing column: ${name}`);
    }

    return index;
  }

  const columns = {
    briefId: col('Brief ID'),
    meetingDate: col('Meeting Date'),
    meetingTime: col('Meeting Time'),
    meetingTitle: col('Meeting Title'),
    calendarEventId: col('Calendar Event ID'),
    briefStatus: col('Brief Status'),
    automationStatus: col('Automation Status'),
    lastRun: col('Last Run')
  };

  for (let rowNumber = 1; rowNumber < queueData.length; rowNumber++) {
    const row = queueData[rowNumber];

    const briefId = String(row[columns.briefId] || '').trim();
    const briefStatus = String(row[columns.briefStatus] || '').trim();

    if (!briefId || briefStatus.toLowerCase() !== 'ready') {
      continue;
    }

    try {
      const meetingDate = row[columns.meetingDate];
      const meetingTime = row[columns.meetingTime];
      const meetingTitle = String(row[columns.meetingTitle] || '').trim();

      const meetingDateTime = buildBriefDateTime(
        meetingDate,
        meetingTime
      );

      const dayStart = new Date(meetingDateTime);
      dayStart.setHours(0, 0, 0, 0);

      const dayEnd = new Date(meetingDateTime);
      dayEnd.setHours(23, 59, 59, 999);

      const calendar = CalendarApp.getDefaultCalendar();

      const events = calendar.getEvents(
        dayStart,
        dayEnd,
        {
          search: meetingTitle
        }
      );

      if (events.length === 0) {
        throw new Error(
          `No calendar event found for "${meetingTitle}".`
        );
      }

      const event = events[0];

      const guestEmails = event
        .getGuestList()
        .map(guest => guest.getEmail())
        .join(', ');

      const sourceRow = [
        briefId,
        'Calendar',
        event.getTitle(),
        event.getId(),
        Utilities.formatDate(
          event.getStartTime(),
          BRIEF_CONFIG.TIME_ZONE,
          'yyyy-MM-dd h:mm a'
        ),
        guestEmails,
        '',
        event.getDescription() || '',
        '',
        '',
        '',
        '',
        'Yes'
      ];

      sourceSheet.appendRow(sourceRow);

      queueSheet
        .getRange(rowNumber + 1, columns.calendarEventId + 1)
        .setValue(event.getId());

      queueSheet
        .getRange(rowNumber + 1, columns.automationStatus + 1)
        .setValue('Calendar Context Added');

      queueSheet
        .getRange(rowNumber + 1, columns.lastRun + 1)
        .setValue(new Date());

    } catch (error) {
      queueSheet
        .getRange(rowNumber + 1, columns.automationStatus + 1)
        .setValue(error.message);
    }
  }
}


function buildBriefDateTime(dateValue, timeValue) {
  let dateText;
  let timeText;

  if (dateValue instanceof Date) {
    dateText = Utilities.formatDate(
      dateValue,
      BRIEF_CONFIG.TIME_ZONE,
      'yyyy-MM-dd'
    );
  } else {
    dateText = String(dateValue).trim();
  }

  if (timeValue instanceof Date) {
    timeText = Utilities.formatDate(
      timeValue,
      BRIEF_CONFIG.TIME_ZONE,
      'h:mm a'
    );
  } else {
    timeText = String(timeValue).trim();
  }

  return Utilities.parseDate(
    `${dateText} ${timeText}`,
    BRIEF_CONFIG.TIME_ZONE,
    'yyyy-MM-dd h:mm a'
  );
}
function pullGmailContext() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = spreadsheet.getSheetByName(BRIEF_CONFIG.QUEUE_SHEET);
  const sourceSheet = spreadsheet.getSheetByName(BRIEF_CONFIG.SOURCE_SHEET);

  if (!queueSheet || !sourceSheet) {
    throw new Error('Required sheet not found.');
  }

  const queueData = queueSheet.getDataRange().getValues();
  const headers = queueData[0];

  function col(name) {
    const index = headers.indexOf(name);

    if (index === -1) {
      throw new Error(`Missing column: ${name}`);
    }

    return index;
  }

  const columns = {
    briefId: col('Brief ID'),
    includeGmail: col('Include Gmail'),
    gmailSearchQuery: col('Gmail Search Query'),
    briefStatus: col('Brief Status'),
    automationStatus: col('Automation Status'),
    lastRun: col('Last Run')
  };

  for (let rowNumber = 1; rowNumber < queueData.length; rowNumber++) {
    const row = queueData[rowNumber];

    const briefId = String(row[columns.briefId] || '').trim();
    const includeGmail = String(row[columns.includeGmail] || '').trim();
    const searchQuery = String(row[columns.gmailSearchQuery] || '').trim();
    const briefStatus = String(row[columns.briefStatus] || '').trim();

    if (
      !briefId ||
      includeGmail.toLowerCase() !== 'yes' ||
      briefStatus.toLowerCase() !== 'ready'
    ) {
      continue;
    }

    if (!searchQuery) {
      queueSheet
        .getRange(rowNumber + 1, columns.automationStatus + 1)
        .setValue('Missing Gmail Search Query');

      continue;
    }

    try {
      const threads = GmailApp.search(searchQuery, 0, 20);

      if (threads.length === 0) {
        throw new Error(
          `No Gmail threads found for query: ${searchQuery}`
        );
      }

      threads.forEach(thread => {
        const messages = thread.getMessages();

        messages.forEach(message => {
          const subject = message.getSubject();
          const from = message.getFrom();
          const date = message.getDate();
          const body = message.getPlainBody();
          const messageId = message.getId();

          const sourceRow = [
            briefId,
            'Gmail',
            subject,
            messageId,
            Utilities.formatDate(
              date,
              BRIEF_CONFIG.TIME_ZONE,
              'yyyy-MM-dd h:mm a'
            ),
            from,
            '',
            body,
            '',
            '',
            '',
            '',
            'Yes'
          ];

          sourceSheet.appendRow(sourceRow);
        });
      });

      queueSheet
        .getRange(rowNumber + 1, columns.automationStatus + 1)
        .setValue('Gmail Context Added');

      queueSheet
        .getRange(rowNumber + 1, columns.lastRun + 1)
        .setValue(new Date());

    } catch (error) {
      queueSheet
        .getRange(rowNumber + 1, columns.automationStatus + 1)
        .setValue(error.message);
    }
  }
}
function pullDriveContext() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = spreadsheet.getSheetByName(BRIEF_CONFIG.QUEUE_SHEET);
  const sourceSheet = spreadsheet.getSheetByName(BRIEF_CONFIG.SOURCE_SHEET);

  if (!queueSheet || !sourceSheet) {
    throw new Error('Required sheet not found.');
  }

  const queueData = queueSheet.getDataRange().getValues();
  const headers = queueData[0];

  function col(name) {
    const index = headers.indexOf(name);

    if (index === -1) {
      throw new Error(`Missing column: ${name}`);
    }

    return index;
  }

  const columns = {
    briefId: col('Brief ID'),
    includeDocs: col('Include Docs'),
    includeSheets: col('Include Sheets'),
    docsSourceUrl: col('Docs Source URL'),
    sheetsSourceUrl: col('Sheets Source URL'),
    sheetsSourceRange: col('Sheets Source Range'),
    briefStatus: col('Brief Status'),
    automationStatus: col('Automation Status'),
    lastRun: col('Last Run')
  };

  for (let rowNumber = 1; rowNumber < queueData.length; rowNumber++) {
    const row = queueData[rowNumber];

    const briefId = String(row[columns.briefId] || '').trim();
    const briefStatus = String(row[columns.briefStatus] || '').trim();

    if (!briefId || briefStatus.toLowerCase() !== 'ready') {
      continue;
    }

    try {
      // GOOGLE DOC SOURCE
      const includeDocs =
        String(row[columns.includeDocs] || '').trim().toLowerCase();

      const docsUrl =
        String(row[columns.docsSourceUrl] || '').trim();

      if (includeDocs === 'yes' && docsUrl) {
        const doc = DocumentApp.openByUrl(docsUrl);
        const docText = doc.getBody().getText();

        sourceSheet.appendRow([
          briefId,
          'Google Doc',
          doc.getName(),
          docsUrl,
          Utilities.formatDate(
            new Date(),
            BRIEF_CONFIG.TIME_ZONE,
            'yyyy-MM-dd h:mm a'
          ),
          '',
          '',
          docText,
          '',
          '',
          '',
          '',
          'Yes'
        ]);
      }

      // GOOGLE SHEET SOURCE
      const includeSheets =
        String(row[columns.includeSheets] || '').trim().toLowerCase();

      const sheetUrl =
        String(row[columns.sheetsSourceUrl] || '').trim();

      const sheetRange =
        String(row[columns.sheetsSourceRange] || '').trim();

      if (includeSheets === 'yes' && sheetUrl && sheetRange) {
        const sourceSpreadsheet = SpreadsheetApp.openByUrl(sheetUrl);

        const rangeParts = sheetRange.split('!');
        const sheetName = rangeParts[0];
        const rangeA1 = rangeParts[1];

        const sourceTab = sourceSpreadsheet.getSheetByName(sheetName);

        if (!sourceTab) {
          throw new Error(`Source tab not found: ${sheetName}`);
        }

        const values = sourceTab.getRange(rangeA1).getDisplayValues();

        const formattedText = values
          .map(rowValues => rowValues.join(' | '))
          .join('\n');

        sourceSheet.appendRow([
          briefId,
          'Google Sheet',
          sourceSpreadsheet.getName(),
          `${sheetUrl}#${sheetRange}`,
          Utilities.formatDate(
            new Date(),
            BRIEF_CONFIG.TIME_ZONE,
            'yyyy-MM-dd h:mm a'
          ),
          '',
          '',
          formattedText,
          '',
          '',
          '',
          '',
          'Yes'
        ]);
      }

      queueSheet
        .getRange(rowNumber + 1, columns.automationStatus + 1)
        .setValue('Drive Context Added');

      queueSheet
        .getRange(rowNumber + 1, columns.lastRun + 1)
        .setValue(new Date());

    } catch (error) {
      queueSheet
        .getRange(rowNumber + 1, columns.automationStatus + 1)
        .setValue(error.message);
    }
  }
}
function generateBriefSourcePacket() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = spreadsheet.getSheetByName(BRIEF_CONFIG.QUEUE_SHEET);
  const sourceSheet = spreadsheet.getSheetByName(BRIEF_CONFIG.SOURCE_SHEET);

  if (!queueSheet || !sourceSheet) {
    throw new Error('Required sheet not found.');
  }

  const queueData = queueSheet.getDataRange().getValues();
  const sourceData = sourceSheet.getDataRange().getValues();

  const queueHeaders = queueData[0];
  const sourceHeaders = sourceData[0];

  function queueCol(name) {
    const index = queueHeaders.indexOf(name);

    if (index === -1) {
      throw new Error(`Missing queue column: ${name}`);
    }

    return index;
  }

  function sourceCol(name) {
    const index = sourceHeaders.indexOf(name);

    if (index === -1) {
      throw new Error(`Missing source column: ${name}`);
    }

    return index;
  }

  const queueColumns = {
    briefId: queueCol('Brief ID'),
    meetingDate: queueCol('Meeting Date'),
    meetingTime: queueCol('Meeting Time'),
    meetingTitle: queueCol('Meeting Title'),
    executive: queueCol('Executive / Owner'),
    briefStatus: queueCol('Brief Status'),
    briefDocUrl: queueCol('Brief Doc URL'),
    automationStatus: queueCol('Automation Status'),
    lastRun: queueCol('Last Run')
  };

  const sourceColumns = {
    briefId: sourceCol('Brief ID'),
    sourceType: sourceCol('Source Type'),
    sourceName: sourceCol('Source Name / Subject'),
    sourceUrl: sourceCol('Source URL / ID'),
    timestamp: sourceCol('Date / Timestamp'),
    owner: sourceCol('Owner / Sender'),
    summary: sourceCol('Summary'),
    decision: sourceCol('Decision Needed'),
    risk: sourceCol('Risk / Blocker'),
    deadline: sourceCol('Deadline'),
    followUp: sourceCol('Follow-Up'),
    included: sourceCol('Included in Brief')
  };

  for (let rowNumber = 1; rowNumber < queueData.length; rowNumber++) {
    const row = queueData[rowNumber];

    const briefId =
      String(row[queueColumns.briefId] || '').trim();

    const briefStatus =
      String(row[queueColumns.briefStatus] || '').trim().toLowerCase();

    if (!briefId || briefStatus !== 'ready') {
      continue;
    }

    try {
      const matchingSources = sourceData.slice(1).filter(sourceRow => {
        const sourceBriefId =
          String(sourceRow[sourceColumns.briefId] || '').trim();

        const included =
          String(sourceRow[sourceColumns.included] || '').trim().toLowerCase();

        return sourceBriefId === briefId && included === 'yes';
      });

      if (matchingSources.length === 0) {
        throw new Error(`No included sources found for ${briefId}.`);
      }

      const meetingTitle =
        String(row[queueColumns.meetingTitle] || '').trim();

      const executive =
        String(row[queueColumns.executive] || '').trim();

      const meetingDate =
        formatBriefValue(row[queueColumns.meetingDate]);

      const meetingTime =
        formatBriefValue(row[queueColumns.meetingTime]);

      const documentTitle =
        `${briefId} — ${meetingTitle} — Source Packet`;

      const doc = DocumentApp.create(documentTitle);
      const body = doc.getBody();

      body.appendParagraph('EXECUTIVE MEETING PREP')
        .setHeading(DocumentApp.ParagraphHeading.TITLE);

      body.appendParagraph(meetingTitle)
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);

      body.appendParagraph(`Brief ID: ${briefId}`);
      body.appendParagraph(`Executive / Owner: ${executive}`);
      body.appendParagraph(`Meeting Date: ${meetingDate}`);
      body.appendParagraph(`Meeting Time: ${meetingTime}`);

      body.appendHorizontalRule();

      body.appendParagraph('SOURCE MATERIAL')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);

      matchingSources.forEach((sourceRow, index) => {
        const type =
          String(sourceRow[sourceColumns.sourceType] || '').trim();

        const name =
          String(sourceRow[sourceColumns.sourceName] || '').trim();

        const timestamp =
          String(sourceRow[sourceColumns.timestamp] || '').trim();

        const owner =
          String(sourceRow[sourceColumns.owner] || '').trim();

        const summary =
          String(sourceRow[sourceColumns.summary] || '').trim();

        body.appendParagraph(
          `${index + 1}. ${type} — ${name}`
        ).setHeading(DocumentApp.ParagraphHeading.HEADING2);

        if (timestamp) {
          body.appendParagraph(`Date / Time: ${timestamp}`);
        }

        if (owner) {
          body.appendParagraph(`Owner / Sender: ${owner}`);
        }

        if (summary) {
          body.appendParagraph(summary);
        }

        body.appendParagraph('');
      });

      body.appendHorizontalRule();

      body.appendParagraph('DECISIONS NEEDED')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);

      appendSourceField(
        body,
        matchingSources,
        sourceColumns.decision,
        sourceColumns.sourceName
      );

      body.appendParagraph('RISKS / BLOCKERS')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);

      appendSourceField(
        body,
        matchingSources,
        sourceColumns.risk,
        sourceColumns.sourceName
      );

      body.appendParagraph('DEADLINES')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);

      appendSourceField(
        body,
        matchingSources,
        sourceColumns.deadline,
        sourceColumns.sourceName
      );

      body.appendParagraph('FOLLOW-UP')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);

      appendSourceField(
        body,
        matchingSources,
        sourceColumns.followUp,
        sourceColumns.sourceName
      );

      doc.saveAndClose();

      queueSheet
        .getRange(rowNumber + 1, queueColumns.briefDocUrl + 1)
        .setValue(doc.getUrl());

      queueSheet
        .getRange(rowNumber + 1, queueColumns.automationStatus + 1)
        .setValue('Source Packet Created');

      queueSheet
        .getRange(rowNumber + 1, queueColumns.lastRun + 1)
        .setValue(new Date());

    } catch (error) {
      queueSheet
        .getRange(rowNumber + 1, queueColumns.automationStatus + 1)
        .setValue(error.message);
    }
  }
}


function appendSourceField(body, sources, fieldIndex, nameIndex) {
  let found = false;

  sources.forEach(sourceRow => {
    const value = String(sourceRow[fieldIndex] || '').trim();

    if (value) {
      const sourceName =
        String(sourceRow[nameIndex] || '').trim();

      body.appendListItem(
        `${value}${sourceName ? ` — Source: ${sourceName}` : ''}`
      );

      found = true;
    }
  });

  if (!found) {
    body.appendParagraph('No items classified yet.');
  }
}


function formatBriefValue(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(
      value,
      BRIEF_CONFIG.TIME_ZONE,
      'yyyy-MM-dd h:mm a'
    );
  }

  return String(value || '').trim();
}



function testGeminiConnection() {
  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty('GEMINI_API_KEY');

  if (!apiKey) {
    throw new Error('Gemini API key not found.');
  }

  const response = UrlFetchApp.fetch(
    'https://generativelanguage.googleapis.com/v1beta/models',
    {
      method: 'get',
      headers: {
        'x-goog-api-key': apiKey
      },
      muteHttpExceptions: true
    }
  );

  const status = response.getResponseCode();

  if (status !== 200) {
    throw new Error(
      'Gemini connection failed. HTTP status: ' + status
    );
  }

  const result = JSON.parse(response.getContentText());

  const models = (result.models || [])
    .filter(model =>
      (model.supportedGenerationMethods || [])
        .includes('generateContent')
    )
    .map(model => model.name);

  Logger.log('Gemini connection successful.');
  Logger.log('Available models: ' + models.join(', '));
}

function testGeminiClassification() {
  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty('GEMINI_API_KEY');

  if (!apiKey) {
    throw new Error('Gemini API key not found.');
  }

  const prompt = `
Read this executive operations update:

The vendor agreement expires September 18.
Legal has not completed its review.
Leadership must decide whether to extend the contract.
The operations team needs to follow up with Legal.

Return only a JSON object with these four fields:
decisionNeeded, riskBlocker, deadline, followUp.

Use only information provided in the update.
`;

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent';

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-goog-api-key': apiKey
    },
    payload: JSON.stringify({
      contents: [
        {
          parts: [{ text: prompt }]
        }
      ],
      generationConfig: {
        responseMimeType: 'application/json'
      }
    }),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();

  if (status !== 200) {
    throw new Error(
      'Gemini API error ' + status + ': ' +
      response.getContentText()
    );
  }

  const result = JSON.parse(response.getContentText());
  const answer = result.candidates[0].content.parts[0].text;

  Logger.log('Gemini classification result:');
  Logger.log(answer);
}


function classifyRemainingBriefSources() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('Source Log');

  if (!sheet) {
    throw new Error('Source Log sheet not found.');
  }

  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  function col(name) {
    const index = headers.indexOf(name);

    if (index === -1) {
      throw new Error('Missing column: ' + name);
    }

    return index;
  }

  const c = {
    briefId: col('Brief ID'),
    type: col('Source Type'),
    summary: col('Summary'),
    decision: col('Decision Needed'),
    risk: col('Risk / Blocker'),
    deadline: col('Deadline'),
    followUp: col('Follow-Up'),
    included: col('Included in Brief')
  };

  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty('GEMINI_API_KEY');

  if (!apiKey) {
    throw new Error('Gemini API key not found.');
  }

  const url =
    'https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent';

  // Find every brief marked Ready in the Briefing Queue.
  const queueData = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName('Briefing Queue')
    .getDataRange()
    .getValues();

  const qId = queueData[0].indexOf('Brief ID');
  const qStatus = queueData[0].indexOf('Brief Status');

  const readyBriefs = new Set(
    queueData.slice(1)
      .filter(r => String(r[qStatus]).trim().toLowerCase() === 'ready')
      .map(r => String(r[qId]).trim())
  );
  let completed = 0;

  // Process only the remaining test sources.
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const sheetRow = i + 1;

    const briefId = String(row[c.briefId] || '').trim();
    const type = String(row[c.type] || '').trim();
    const included = String(row[c.included] || '')
      .trim().toLowerCase();

    const sourceText = String(row[c.summary] || '').trim();

    if (
            !readyBriefs.has(briefId) ||
      included !== 'yes' ||
      !sourceText
    ) {
      continue;
    }

    const existingDecision =
      String(row[c.decision] || '').trim();

    const existingRisk =
      String(row[c.risk] || '').trim();

    const existingDeadline =
      String(row[c.deadline] || '').trim();

    const existingFollowUp =
      String(row[c.followUp] || '').trim();

    // Preserve results from any previously completed row.
    if (
      existingDecision ||
      existingDeadline ||
      existingFollowUp ||
      (existingRisk &&
        !existingRisk.startsWith('AI Error:'))
    ) {
      Logger.log('Skipping previously classified row ' + sheetRow);
      continue;
    }

    const prompt = `
Classify this executive operations source material.

SOURCE:
${sourceText}

Return only a JSON object with these four fields:
decisionNeeded, riskBlocker, deadline, followUp.

Rules:
- Use only information supported by the source.
- Do not invent missing details.
- Keep each field concise.
- If no information is available for a field, use an empty string.
`;

    try {
      let answer = null;

      // Retry temporary API failures up to three times.
      for (let attempt = 1; attempt <= 3; attempt++) {
        const response = UrlFetchApp.fetch(url, {
          method: 'post',
          contentType: 'application/json',
          headers: {
            'x-goog-api-key': apiKey
          },
          payload: JSON.stringify({
            contents: [
              {
                parts: [{ text: prompt }]
              }
            ],
            generationConfig: {
              responseMimeType: 'application/json'
            }
          }),
          muteHttpExceptions: true
        });

        const status = response.getResponseCode();

        if (status === 200) {
          const result = JSON.parse(
            response.getContentText()
          );

          answer = JSON.parse(
            result.candidates[0].content.parts[0].text
          );

          break;
        }

        if (
          [429, 500, 502, 503, 504].includes(status) &&
          attempt < 3
        ) {
          Utilities.sleep(attempt * 5000);
          continue;
        }

        throw new Error(
          'Gemini API error ' + status + ': ' +
          response.getContentText()
        );
      }

      if (!answer) {
        throw new Error('No classification returned.');
      }

      // Write all four results to the correct source row.
      sheet.getRange(sheetRow, c.decision + 1, 1, 4)
        .setValues([[
          answer.decisionNeeded || '',
          answer.riskBlocker || '',
          answer.deadline || '',
          answer.followUp || ''
        ]]);

      completed++;

      Logger.log(
        'Classification completed for ' +
        type + ' — row ' + sheetRow
      );

    } catch (error) {
      // Keep existing source information unchanged.
      Logger.log(
        'Classification failed for row ' +
        sheetRow + ': ' + error.message
      );
    }
  }

  Logger.log(
    'Total sources classified: ' + completed
  );
}
function generateExecutiveBrief() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const queueSheet = spreadsheet.getSheetByName('Briefing Queue');
  const sourceSheet = spreadsheet.getSheetByName('Source Log');

  if (!queueSheet || !sourceSheet) {
    throw new Error('Required sheet not found.');
  }

  // Read the briefing instructions and classified sources.
  const queueData = queueSheet.getDataRange().getValues();
  const sourceData = sourceSheet.getDataRange().getValues();
  const queueHeaders = queueData[0];
  const sourceHeaders = sourceData[0];

  function queueCol(name) {
    const index = queueHeaders.indexOf(name);
    if (index === -1) {
      throw new Error('Missing queue column: ' + name);
    }
    return index;
  }

  function sourceCol(name) {
    const index = sourceHeaders.indexOf(name);
    if (index === -1) {
      throw new Error('Missing source column: ' + name);
    }
    return index;
  }

  const q = {
    briefId: queueCol('Brief ID'),
    meetingTitle: queueCol('Meeting Title'),
    meetingDate: queueCol('Meeting Date'),
    meetingTime: queueCol('Meeting Time'),
    executive: queueCol('Executive / Owner'),
    briefStatus: queueCol('Brief Status'),
    status: queueCol('Automation Status'),
    lastRun: queueCol('Last Run')
  };

  const s = {
    briefId: sourceCol('Brief ID'),
    type: sourceCol('Source Type'),
    name: sourceCol('Source Name / Subject'),
    url: sourceCol('Source URL / ID'),
    decision: sourceCol('Decision Needed'),
    risk: sourceCol('Risk / Blocker'),
    deadline: sourceCol('Deadline'),
    followUp: sourceCol('Follow-Up'),
    included: sourceCol('Included in Brief')
  };

  // Create a separate link column for the finished brief.
  let finalUrlCol = queueHeaders.indexOf('Final Brief URL');

  if (finalUrlCol === -1) {
    finalUrlCol = queueHeaders.length;
    queueSheet.getRange(1, finalUrlCol + 1).setValue('Final Brief URL');
  }

  let created = 0;

  // Build one brief for every Ready row in the queue.
  for (let i = 1; i < queueData.length; i++) {
    const queueRow = queueData[i];
    const sheetRow = i + 1;

    const briefId = String(queueRow[q.briefId] || '').trim();
    const briefStatus =
      String(queueRow[q.briefStatus] || '').trim().toLowerCase();

    if (!briefId || briefStatus !== 'ready') {
      continue;
    }

    // Prevent duplicate documents when the function runs again.
    const existingUrl = String(queueRow[finalUrlCol] || '').trim();

    if (existingUrl) {
      Logger.log(briefId + ' already has a brief: ' + existingUrl);
      continue;
    }

    try {
      const sources = sourceData.slice(1).filter(row =>
        String(row[s.briefId]).trim() === briefId &&
        String(row[s.included]).trim().toLowerCase() === 'yes'
      );

      if (sources.length === 0) {
        throw new Error('No included sources found for ' + briefId + '.');
      }

      const doc = createBriefDoc(briefId, queueRow, q, sources, s);

      // Record the finished brief and close the row.
      queueSheet.getRange(sheetRow, finalUrlCol + 1).setValue(doc.getUrl());
      queueSheet.getRange(sheetRow, q.status + 1).setValue('Executive Brief Created');
      queueSheet.getRange(sheetRow, q.lastRun + 1).setValue(new Date());
      queueSheet.getRange(sheetRow, q.briefStatus + 1).setValue('Complete');

      created++;
      Logger.log('Executive Brief created for ' + briefId + ': ' + doc.getUrl());

    } catch (error) {
      queueSheet.getRange(sheetRow, q.status + 1).setValue(error.message);
    }
  }

  Logger.log('Executive briefs created: ' + created);
}


function createBriefDoc(briefId, queueRow, q, sources, s) {
  const meetingTitle = String(queueRow[q.meetingTitle] || '').trim();
  const executive = String(queueRow[q.executive] || '').trim();

  const doc = DocumentApp.create(
    briefId + ' — ' + meetingTitle + ' — Executive Brief'
  );

  const body = doc.getBody();

  body.appendParagraph('EXECUTIVE MEETING BRIEF')
    .setHeading(DocumentApp.ParagraphHeading.TITLE);

  body.appendParagraph(meetingTitle)
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  body.appendParagraph('Brief ID: ' + briefId);
  body.appendParagraph('Executive / Owner: ' + executive);
  body.appendParagraph('Meeting Date: ' + formatBriefValue(queueRow[q.meetingDate]));
  body.appendParagraph('Meeting Time: ' + formatBriefValue(queueRow[q.meetingTime]));
  body.appendParagraph('Sources reviewed: ' + sources.length);

  body.appendHorizontalRule();

  addBriefSection(body, sources, 'DECISIONS NEEDED', s.decision, s.name);
  addBriefSection(body, sources, 'RISKS / BLOCKERS', s.risk, s.name);
  addBriefSection(body, sources, 'DEADLINES', s.deadline, s.name);
  addBriefSection(body, sources, 'FOLLOW-UP ACTIONS', s.followUp, s.name);

  // Preserve references to the original information.
  body.appendParagraph('SOURCE REFERENCES')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  sources.forEach((row, index) => {
    const type = String(row[s.type] || '').trim();
    const name = String(row[s.name] || '').trim();
    const reference = String(row[s.url] || '').trim();

    body.appendParagraph((index + 1) + '. ' + type + ' — ' + name);

    if (reference) {
      body.appendParagraph('Reference: ' + reference);
    }
  });

  body.appendHorizontalRule();

  body.appendParagraph(
    'AI-generated classifications require human review ' +
    'before executive distribution or action.'
  );

  doc.saveAndClose();
  return doc;
}


function addBriefSection(body, sources, title, fieldIndex, nameIndex) {
  body.appendParagraph(title)
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  let found = false;
  const seen = new Set();

  sources.forEach(row => {
    const rawValue = row[fieldIndex];

    const value = rawValue instanceof Date
      ? Utilities.formatDate(rawValue, BRIEF_CONFIG.TIME_ZONE, 'MMMM d, yyyy')
      : String(rawValue || '').trim();

    if (!value || value.startsWith('AI Error:')) {
      return;
    }

    if (seen.has(value.toLowerCase())) {
      return;
    }

    seen.add(value.toLowerCase());

    body.appendListItem(value);
    body.appendParagraph('Source: ' + String(row[nameIndex] || '').trim());

    found = true;
  });

  if (!found) {
    body.appendParagraph('No items identified in the classified sources.');
  }
}


