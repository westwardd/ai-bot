class RealEstateAssistant {
  constructor(sheetId, openAiKey) {
    this.SHEET_ID = sheetId;
    this.OPENAI_API_KEY = openAiKey;
    this.MAX_EMAILS_PER_RUN = 5;       // максимум писем за запуск
    this.MAX_EXECUTION_TIME_MS = 5 * 60 * 1000; // максимум 5 минут

    const ss = SpreadsheetApp.openById(this.SHEET_ID);
    this.sheetClients = ss.getSheetByName('Clients');
    this.sheetOwners = ss.getSheetByName('Owners');
    this.sheetViewings = ss.getSheetByName('Viewings');
  }

  run() {
    const startTime = Date.now();
    let processed = 0;

    while (processed < this.MAX_EMAILS_PER_RUN && (Date.now() - startTime) < this.MAX_EXECUTION_TIME_MS) {
      const threads = GmailApp.search('is:unread', 0, 1);
      if (threads.length === 0) {
        Logger.log('No unread emails left to process.');
        break;
      }

      try {
        const thread = threads[0];
        const messages = thread.getMessages();
        const lastMessage = messages[messages.length - 1];
        const sender = this.getEmailOnly(lastMessage.getFrom());
        const fullConversation = messages.map(m => m.getPlainBody()).join('\n---\n');

        const status = this.getStatusByEmail(sender);
        const prompt = this.generatePrompt(fullConversation, status);
        const gptResponse = this.askGPT(prompt);

        if (!gptResponse) {
          thread.markRead();
          processed++;
          continue;
        }

        const { role, data, reply } = gptResponse;

        if (role === 'client') {
          this.handleClient(sender, data || {}, reply || '', lastMessage);
        } else if (role === 'owner') {
          this.handleOwner(sender, data || {}, reply || '', lastMessage);
        } else {
          lastMessage.reply(reply || 'Thank you for your message.');
        }

        thread.markRead();
      } catch (e) {
        Logger.log('Error processing email: ' + e);
        try {
          threads[0].markRead();
        } catch (_) {}
      }

      processed++;
    }

    this.finalizeViewingsFromSheet();
  }

  handleClient(sender, data, reply, message) {
    let status = this.getStatusByEmail(sender);
    const rowExists = this.findRowByEmail(this.sheetClients, sender);

    if (!rowExists) {
      this.sheetClients.appendRow([new Date(), sender, data.location || '', data.budget || '', data.type || '', 'new']);
      status = 'new';
      Logger.log(`Client added: ${sender}`);
    }

    if (!data.location || !data.budget || !data.type) {
      message.reply(reply || 'Please provide your desired location, budget, and property type.');
      this.updateStatusByEmail('Clients', sender, 'awaiting_details');
      return;
    }

    Logger.log(`Searching listings for client ${sender} with location=${data.location}, budget=${data.budget}, type=${data.type}`);
    const matches = this.findMatchingListings(data.location, data.budget, data.type);
    Logger.log(`Matches found: ${matches.length}`);

    if (matches.length === 0) {
      message.reply('We will notify you when matching properties are available.');
      this.updateStatusByEmail('Clients', sender, 'no_matches');
      return;
    }

    if (['new', 'awaiting_details', 'no_matches'].includes(status)) {
      const listingsText = matches.map((m, i) => `${i + 1}) ${m.description} — ${m.location}, $${m.price}`).join('\n');
      const matchReply = `We found ${matches.length} properties:\n${listingsText}\nPlease reply with a preferred viewing time.`;
      message.reply(matchReply);
      this.updateStatusByEmail('Clients', sender, 'waiting_for_time');

      matches.forEach(match => {
        this.logViewing(sender, match.ownerEmail);
        GmailApp.sendEmail(match.ownerEmail,
          'Viewing Request',
          `Client ${sender} is interested in your property at ${match.location}.\nPlease confirm if you are available for a viewing.`);
        this.updateStatusByEmail('Owners', match.ownerEmail, 'awaiting_confirmation');
      });
      return;
    }

    if (status === 'waiting_for_time' && data.viewing_time) {
      this.updateViewingTimeWithChecks(sender, data.viewing_time);
      const owners = this.getOwnersByClient(sender);
      owners.forEach(owner => {
        GmailApp.sendEmail(owner,
          'Viewing Time Proposed',
          `Client ${sender} proposes: ${data.viewing_time}.\nPlease confirm or suggest another time.`);
      });

      this.updateStatusByEmail('Clients', sender, 'waiting_owner_response');
      message.reply('Your time was sent to the owner. Awaiting confirmation.');
      return;
    }

    if (status === 'waiting_for_time') {
      message.reply('Please provide your preferred time for the viewing.');
      return;
    }

    message.reply(reply || 'Thank you for your message.');
  }

  handleOwner(sender, data, reply, message) {
    let status = this.getStatusByEmail(sender);
    const rowExists = this.findRowByEmail(this.sheetOwners, sender);

    if (!rowExists) {
      this.sheetOwners.appendRow([new Date(), sender, data.location || '', data.price || '', data.description || '', 'new']);
      status = 'new';
    }

    const decisionRaw = data.confirmation_time || '';
    const decline = data.decline === true || data.decline === 'true' || data.decline === 'yes';

    if (status === 'awaiting_confirmation' && (decisionRaw || decline)) {
      const decision = decline ? 'declined' : 'confirmed';
      const affectedClients = this.updateViewingStatusByOwner(sender, decision);
      if (affectedClients.length > 0) {
        affectedClients.forEach(clientEmail => {
          if (decision === 'confirmed') {
            GmailApp.sendEmail(clientEmail, 'Viewing Confirmed', `The owner has confirmed your proposed time.`);
            this.updateStatusByEmail('Clients', clientEmail, 'meeting_confirmed');
          } else {
            GmailApp.sendEmail(clientEmail, 'Viewing Declined', `The owner is not available at that time. Please suggest another.`);
            this.updateStatusByEmail('Clients', clientEmail, 'awaiting_new_time');
          }
        });
        this.updateStatusByEmail('Owners', sender, decision);
        message.reply(`Thanks! We've updated the client(s).`);
        return;
      } else {
        message.reply(`No matching viewing found to confirm.`);
        return;
      }
    }

    if (!data.location || !data.price || !data.description) {
      message.reply(reply || 'Please provide location, price, and description.');
      this.updateStatusByEmail('Owners', sender, 'awaiting_details');
      return;
    }

    const matches = this.findMatchingClients(data.location, data.price, data.description);

    if (matches.length === 0) {
      message.reply('Currently no interested clients. We will notify you when they appear.');
      this.updateStatusByEmail('Owners', sender, 'no_clients');
      return;
    }

    const clientsList = matches.map((c, i) => `${i + 1}) ${c.email} — ${c.location}, budget up to $${c.budget}`).join('\n');
    message.reply(`We found ${matches.length} potential clients:\n${clientsList}`);
    this.updateStatusByEmail('Owners', sender, 'notified');
  }

  getEmailOnly(str) {
    if (!str) return '';
    const match = str.match(/<(.+?)>/);
    return match ? match[1].trim().toLowerCase() : str.trim().toLowerCase();
  }

  findRowByEmail(sheet, email) {
    const data = sheet.getDataRange().getValues();
    const cleanEmail = this.getEmailOnly(email);
    for (let i = 1; i < data.length; i++) {
      const sheetEmail = this.getEmailOnly(data[i][1]);
      if (sheetEmail === cleanEmail) return i + 1;
    }
    return null;
  }

  updateStatusByEmail(sheetName, email, status) {
    const sheet = SpreadsheetApp.openById(this.SHEET_ID).getSheetByName(sheetName);
    const row = this.findRowByEmail(sheet, email);
    if (row) {
      sheet.getRange(row, 6).setValue(status);
      Logger.log(`Status updated: ${sheetName} - ${email} -> ${status}`);
    }
  }

  getStatusByEmail(email) {
    const ss = SpreadsheetApp.openById(this.SHEET_ID);
    const sheets = ['Clients', 'Owners'];
    for (const name of sheets) {
      const sheet = ss.getSheetByName(name);
      const row = this.findRowByEmail(sheet, email);
      if (row) return sheet.getRange(row, 6).getValue();
    }
    return '';
  }

  generatePrompt(convo, status) {
    return `
You are a real estate assistant analyzing the following conversation:
${convo}

Based on this message, extract JSON with the following fields:
{
  "role": "client" or "owner",
  "data": {
    "location": "...",
    "budget": "...",
    "type": "...",
    "price": "...",
    "description": "...",
    "viewing_time": "...",
    "confirmation_time": "...",
    "decline": true or false
  },
  "reply": "short polite English response"
}

- If the owner confirms a viewing time, put that time in "confirmation_time".
- If the owner declines the viewing, set "decline" to true.
- Otherwise, keep these fields empty or false.
Return ONLY JSON.
`;
  }

  askGPT(prompt) {
    try {
      const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: `Bearer ${this.OPENAI_API_KEY}` },
        payload: JSON.stringify({
          model: 'gpt-4o',
          messages: [{ role: 'user', content: prompt }],
          temperature: 0.3
        })
      });
      const json = JSON.parse(res.getContentText());
      let content = json.choices?.[0]?.message?.content || '';
      content = content.replace(/```json/i, '').replace(/```/g, '').trim();
      return JSON.parse(content);
    } catch (e) {
      Logger.log('GPT Error: ' + e);
      return null;
    }
  }

  findMatchingListings(location, budget, type) {
    const rows = this.sheetOwners.getDataRange().getValues();
    const matches = [];
    const budgetNum = parseInt((budget || '').toString().replace(/\D/g, ''), 10) || 0;

    for (let i = 1; i < rows.length; i++) {
      const rowLoc = (rows[i][2] || '').toLowerCase();
      const price = parseInt((rows[i][3] || '').toString().replace(/\D/g, ''), 10);
      const desc = (rows[i][4] || '').toLowerCase();

      if (
        (!location || rowLoc.includes(location.toLowerCase())) &&
        (!type || desc.includes(type.toLowerCase())) &&
        (price <= budgetNum || !budgetNum)
      ) {
        matches.push({
          location: rows[i][2],
          price: price,
          description: rows[i][4],
          ownerEmail: this.getEmailOnly(rows[i][1])
        });
      }
    }
    return matches;
  }

  findMatchingClients(location, price, description) {
    const rows = this.sheetClients.getDataRange().getValues();
    const matches = [];
    const priceNum = parseInt((price || '').toString().replace(/\D/g, ''), 10);

    for (let i = 1; i < rows.length; i++) {
      const clientLoc = (rows[i][2] || '').toLowerCase();
      const clientBudget = parseInt((rows[i][3] || '').toString().replace(/\D/g, ''), 10);
      const clientType = (rows[i][4] || '').toLowerCase();

      if (
        (!location || clientLoc.includes(location.toLowerCase())) &&
        (!description || clientType.includes(description.toLowerCase())) &&
        (clientBudget >= priceNum || !priceNum)
      ) {
        matches.push({
          email: this.getEmailOnly(rows[i][1]),
          location: rows[i][2],
          budget: clientBudget,
          type: rows[i][4]
        });
      }
    }
    return matches;
  }

  logViewing(clientEmail, ownerEmail) {
    this.sheetViewings.appendRow([new Date(), clientEmail, ownerEmail, '', 'pending']);
  }

  updateViewingTimeWithChecks(clientEmail, time) {
    const viewings = this.sheetViewings.getDataRange().getValues();
    const ownersToNotify = new Set();

    for (let i = 1; i < viewings.length; i++) {
      if (this.getEmailOnly(viewings[i][1]) === this.getEmailOnly(clientEmail)) {
        const rowIdx = i + 1;
        this.sheetViewings.getRange(rowIdx, 4).setValue(time);
        this.sheetViewings.getRange(rowIdx, 5).setValue('scheduled');
      }
    }

    // Проверяем подтверждены ли все просмотры
    let allConfirmed = true;
    for (let i = 1; i < viewings.length; i++) {
      if (this.getEmailOnly(viewings[i][1]) === this.getEmailOnly(clientEmail)) {
        const status = (viewings[i][4] || '').toLowerCase();
        ownersToNotify.add(this.getEmailOnly(viewings[i][2]));
        if (status !== 'confirmed') {
          allConfirmed = false;
        }
      }
    }

    if (allConfirmed && ownersToNotify.size > 0) {
      this.updateStatusByEmail('Clients', clientEmail, 'meeting_confirmed');
      ownersToNotify.forEach(owner => this.updateStatusByEmail('Owners', owner, 'confirmed'));
    }
  }

  updateViewingStatusByOwner(ownerEmail, decision) {
    const viewings = this.sheetViewings.getDataRange().getValues();
    const affectedClients = [];

    for (let i = 1; i < viewings.length; i++) {
      if (this.getEmailOnly(viewings[i][2]) === this.getEmailOnly(ownerEmail)) {
        const rowIdx = i + 1;
        if (decision === 'confirmed') {
          this.sheetViewings.getRange(rowIdx, 5).setValue('confirmed');
        } else {
          this.sheetViewings.getRange(rowIdx, 5).setValue('declined');
        }
        affectedClients.push(this.getEmailOnly(viewings[i][1]));
      }
    }
    return [...new Set(affectedClients)];
  }

  getOwnersByClient(clientEmail) {
    const viewings = this.sheetViewings.getDataRange().getValues();
    const owners = new Set();

    for (let i = 1; i < viewings.length; i++) {
      if (this.getEmailOnly(viewings[i][1]) === this.getEmailOnly(clientEmail)) {
        owners.add(this.getEmailOnly(viewings[i][2]));
      }
    }
    return [...owners];
  }

  finalizeViewingsFromSheet() {
    const data = this.sheetViewings.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const client = this.getEmailOnly(data[i][1]);
      const owner = this.getEmailOnly(data[i][2]);
      const time = data[i][3];
      const status = (data[i][4] || '').toLowerCase();

      if (!time || !status) continue;

      if (status === 'confirmed') {
        this.updateStatusByEmail('Clients', client, 'meeting_confirmed');
        this.updateStatusByEmail('Owners', owner, 'meeting_confirmed');

        GmailApp.sendEmail(client, 'Viewing Confirmed', `Your viewing with the owner is confirmed for: ${time}.`);
        GmailApp.sendEmail(owner, 'Viewing Confirmed', `Viewing with client ${client} is confirmed for: ${time}.`);

        this.sheetViewings.getRange(i + 1, 5).setValue('finalized');
      }

      if (status === 'declined') {
        this.updateStatusByEmail('Clients', client, 'awaiting_new_time');
        this.updateStatusByEmail('Owners', owner, 'declined');

        GmailApp.sendEmail(client, 'Viewing Declined', `The owner declined your proposed time. Please suggest another.`);
        this.sheetViewings.getRange(i + 1, 5).setValue('declined_final');
      }
    }
  }
}

// === ЗАПУСК ===
// Подставь свои значения в переменные ниже
const SHEET_ID = '1FBoyUucC3zPj5v7zVzf9ut4Pk8pEAb9rjjyuY7651ew';
const OPENAI_API_KEY = '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';

function main() {
  const assistant = new RealEstateAssistant(SHEET_ID, OPENAI_API_KEY);
  assistant.run();
}
