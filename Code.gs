// LBV Command - Google Apps Script Backend
// GAS is the single source of truth. The upsert action protects all data.

var SHEET_ID = '1sPVpASsTB6q8P2kkIPJkO4_qKAL7sk7m6dAnBC4iuYE';
var ANTHROPIC_KEY = '[REDACTED_ANTHROPIC_KEY_SEE_LOCAL_COPY]';

function doGet(e) {
  if (e && e.parameter && e.parameter.code) {
    var result = handleOAuthCallback(e.parameter.code);
    if (result === 'success') {
      return HtmlService.createHtmlOutput('<html><body><script>window.opener&&window.opener.postMessage("auth_complete","*");window.close();</script><p>Connected. You can close this window.</p></body></html>');
    }
    return HtmlService.createHtmlOutput('<p>Auth failed: ' + result + '</p>');
  }
  return HtmlService.createHtmlOutputFromFile('Page')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    if (action === 'read') return respond(readSheet(body.sheet));
    if (action === 'write') return respond(writeSheet(body.sheet, body.data));
    if (action === 'upsert') return respond(upsertSheet(body.sheet, body.data));
    if (action === 'delete') return respond(deleteRow(body.sheet, body.id));
    if (action === 'inbox') return respond(fetchInbox(body.token));
    if (action === 'getEmails') return respond(getEmails());
    if (action === 'getSentItems') return respond(getSentItems());
    if (action === 'apollo') return respond(apolloProxy(body.endpoint, body.params));
    if (action === 'intentSignals') return respond(getIntentSignals());
    if (action === 'ai') return respond(aiClassify(body.emails));
    return respond({error: 'Unknown action: ' + action});
  } catch(err) {
    return respond({error: err.toString()});
  }
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function readSheet(sheetName) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    var headers = data[0].map(function(h){ return String(h).trim(); });
    return data.slice(1).map(function(row) {
      var obj = {};
      headers.forEach(function(h, i) { obj[h] = row[i] !== undefined ? row[i] : ''; });
      return obj;
    });
  } catch(err) {
    return {error: err.toString()};
  }
}

function writeSheet(sheetName, data) {
  try {
    if (!data || !data.length) return {ok: true, rows: 0};
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    var headers = Object.keys(data[0]);
    var rows = data.map(function(row) {
      return headers.map(function(h) { return row[h] !== undefined ? row[h] : ''; });
    });
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    return {ok: true, rows: rows.length};
  } catch(err) {
    return {error: err.toString()};
  }
}

function upsertSheet(sheetName, data) {
  try {
    if (!data || !data.length) return {ok: true, updated: 0, added: 0};
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      return writeSheet(sheetName, data);
    }
    var existing = sheet.getDataRange().getValues();
    if (existing.length < 1) return writeSheet(sheetName, data);
    var headers = existing[0].map(function(h){ return String(h).trim(); });
    var idCol = headers.indexOf('id');
    if (idCol < 0) return writeSheet(sheetName, data);
    var idToRow = {};
    for (var i = 1; i < existing.length; i++) {
      var rowId = String(existing[i][idCol]).trim();
      if (rowId) idToRow[rowId] = i + 1;
    }
    var incomingHeaders = data.length ? Object.keys(data[0]) : [];
    incomingHeaders.forEach(function(h) {
      if (headers.indexOf(h) < 0) {
        headers.push(h);
        sheet.getRange(1, headers.length).setValue(h);
      }
    });
    var updated = 0, added = 0;
    data.forEach(function(row) {
      var rowId = String(row.id || '').trim();
      if (!rowId) return;
      var values = headers.map(function(h) { return row[h] !== undefined ? row[h] : ''; });
      if (idToRow[rowId]) {
        sheet.getRange(idToRow[rowId], 1, 1, values.length).setValues([values]);
        updated++;
      } else {
        var lastRow = sheet.getLastRow();
        sheet.getRange(lastRow + 1, 1, 1, values.length).setValues([values]);
        idToRow[rowId] = lastRow + 1;
        added++;
      }
    });
    return {ok: true, updated: updated, added: added};
  } catch(err) {
    return {error: err.toString()};
  }
}

function fetchInbox(token) {
  try {
    if (!token) return {error: 'No token'};
    var url = 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=50&$orderby=receivedDateTime desc&$select=id,subject,from,receivedDateTime,isRead,bodyPreview,webLink';
    var resp = UrlFetchApp.fetch(url, {
      headers: {Authorization: 'Bearer ' + token},
      muteHttpExceptions: true
    });
    var data = JSON.parse(resp.getContentText());
    if (data.error) return {error: data.error.message};
    return data.value || [];
  } catch(err) {
    return {error: err.toString()};
  }
}

function apolloProxy(endpoint, params) {
  try {
    var apiKey = 'tMFC_OULnqIQbqKHCGgsVg';
    if (!endpoint) {
      var url = 'https://api.apollo.io/v1/emailer_campaigns/search';
      var resp = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        headers: {'X-Api-Key': apiKey, 'Cache-Control': 'no-cache'},
        payload: JSON.stringify({per_page: 25}),
        muteHttpExceptions: true
      });
      var raw = resp.getContentText();
      var data = JSON.parse(raw);
      var seqs = data.emailer_campaigns || [];
      seqs = seqs.map(function(seq) {
        try {
          var cResp = UrlFetchApp.fetch('https://api.apollo.io/v1/contacts/search', {
            method: 'post',
            contentType: 'application/json',
            headers: {'X-Api-Key': apiKey, 'Cache-Control': 'no-cache'},
            payload: JSON.stringify({emailer_campaign_ids: [seq.id], per_page: 1}),
            muteHttpExceptions: true
          });
          var cData = JSON.parse(cResp.getContentText());
          seq.active_contact_count = (cData.pagination && cData.pagination.total_entries) || 0;
        } catch(e) {
          seq.active_contact_count = 0;
        }
        seq.num_sent_email_count = seq.num_sent_email_count || seq.unique_delivered || seq.delivered_count || 0;
        seq.num_opened_email_count = seq.num_opened_email_count || seq.unique_opened || seq.open_count || 0;
        seq.num_replied_email_count = seq.num_replied_email_count || seq.unique_replied || seq.reply_count || 0;
        return seq;
      });
      data.emailer_campaigns = seqs;
      return data;
    }
    var url = 'https://api.apollo.io/v1/' + endpoint;
    var resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: {'X-Api-Key': apiKey, 'Cache-Control': 'no-cache'},
      payload: JSON.stringify(params || {}),
      muteHttpExceptions: true
    });
    return JSON.parse(resp.getContentText());
  } catch(err) {
    return {error: err.toString()};
  }
}

function getIntentSignals() {
  try {
    var apiKey = 'tMFC_OULnqIQbqKHCGgsVg';
    var signals = [];
    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 14);
    var cutoffStr = cutoff.toISOString();
    var seqResp = UrlFetchApp.fetch('https://api.apollo.io/v1/emailer_campaigns/search', {
      method: 'post',
      contentType: 'application/json',
      headers: {'X-Api-Key': apiKey, 'Cache-Control': 'no-cache'},
      payload: JSON.stringify({per_page: 25}),
      muteHttpExceptions: true
    });
    var seqData = JSON.parse(seqResp.getContentText());
    var seqs = seqData.emailer_campaigns || [];
    seqs.forEach(function(seq) {
      try {
        var evtResp = UrlFetchApp.fetch('https://api.apollo.io/v1/emailer_messages/search', {
          method: 'post',
          contentType: 'application/json',
          headers: {'X-Api-Key': apiKey, 'Cache-Control': 'no-cache'},
          payload: JSON.stringify({emailer_campaign_id: seq.id, per_page: 200}),
          muteHttpExceptions: true
        });
        var evtData = JSON.parse(evtResp.getContentText());
        var messages = evtData.emailer_messages || [];
        messages.forEach(function(msg) {
          var status = (msg.status || '').toLowerCase();
          var updatedAt = msg.updated_at || msg.completed_at || msg.created_at || '';
          var isRecent = updatedAt >= cutoffStr;
          var contactName = msg.to_name || (msg.contact && msg.contact.name) || msg.to_email || '';
          var contactEmail = msg.to_email || '';
          var company = (msg.contact && msg.contact.organization_name) || '';
          if (status === 'replied' && isRecent) {
            signals.push({signal:'replied',contactName:contactName,contactEmail:contactEmail,company:company,sequenceName:seq.name||'',date:updatedAt});
          } else if (status === 'clicked' && isRecent) {
            signals.push({signal:'clicked',contactName:contactName,contactEmail:contactEmail,company:company,sequenceName:seq.name||'',date:updatedAt});
          } else if (status === 'opened' && isRecent) {
            signals.push({signal:'opened',contactName:contactName,contactEmail:contactEmail,company:company,sequenceName:seq.name||'',date:updatedAt});
          } else if (status === 'delivered' && isRecent) {
            signals.push({signal:'delivered',contactName:contactName,contactEmail:contactEmail,company:company,sequenceName:seq.name||'',date:updatedAt});
          }
        });
      } catch(seqErr) {
        Logger.log('Seq error ' + seq.id + ': ' + seqErr.toString());
      }
    });
    var rank = {replied:4,clicked:3,opened:2,delivered:1};
    signals.sort(function(a,b){return (rank[b.signal]||0)-(rank[a.signal]||0);});
    Logger.log('Intent signals found: ' + signals.length);
    return {signals: signals};
  } catch(err) {
    Logger.log('getIntentSignals error: ' + err.toString());
    return {error: err.toString()};
  }
}

function aiClassify(emails) {
  try {
    if (!emails || !emails.length) return [];
    var results = emails.map(function(e) {
      var subj = (e.subject || '').toLowerCase();
      var sender = (e.senderEmail || '').toLowerCase();
      var summary = (e.summary || '').toLowerCase();
      var noiseDomains = ['linkedin.com','apollo.io','mail.apollo.io','fathom.video',
        'sevenrooms.com','moneysupermarketmail.com','creditsafeuk.com',
        'aceledlight.com','cooco.cc','mason-led.com','alcanside.com',
        'clarionevents.com','upshine.com'];
      var isNoise = noiseDomains.some(function(d){ return sender.indexOf(d) >= 0; });
      if (isNoise) return {id: e.id, urgency: 'info', isProspect: false};
      var hotWords = ['invoice','payment','urgent','asap','overdue','credit','quote accepted',
        'order','proforma','po number','purchase order'];
      var isHot = hotWords.some(function(w){ return subj.indexOf(w) >= 0 || summary.indexOf(w) >= 0; });
      if (isHot) return {id: e.id, urgency: 'hot', isProspect: true};
      var warmWords = ['follow up','meeting','call','proposal','visit','interested',
        'availability','schedule','confirm','update','reply','response needed'];
      var isWarm = warmWords.some(function(w){ return subj.indexOf(w) >= 0 || summary.indexOf(w) >= 0; });
      if (isWarm) return {id: e.id, urgency: 'warm', isProspect: false};
      return {id: e.id, urgency: 'info', isProspect: false};
    });
    try {
      var uncertain = emails.filter(function(e, i){ return results[i].urgency === 'info'; });
      if (uncertain.length > 0) {
        var prompt = 'Classify these emails for a commercial LED lighting sales director. ' +
          'Return JSON array with {id, urgency: hot|warm|info, isProspect: true|false}. ' +
          'Hot = needs action today. Warm = prospect/client engaging. Info = noise/newsletter/admin.\n\n' +
          JSON.stringify(uncertain.map(function(e){
            return {id: e.id, subject: e.subject, sender: e.sender, summary: (e.summary||'').slice(0,150)};
          }));
        var resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
          method: 'post',
          contentType: 'application/json',
          headers: {'x-api-key': ANTHROPIC_KEY, 'anthropic-version': '2023-06-01'},
          payload: JSON.stringify({
            model: 'claude-haiku-4-5-20251001',
            max_tokens: 1000,
            messages: [{role: 'user', content: prompt + '\nReturn only valid JSON array.'}]
          }),
          muteHttpExceptions: true
        });
        var aiData = JSON.parse(resp.getContentText());
        if (aiData.content && aiData.content[0]) {
          var aiResults = JSON.parse(aiData.content[0].text.replace(/```json|```/g,'').trim());
          var aiMap = {};
          aiResults.forEach(function(r){ aiMap[r.id] = r; });
          results = results.map(function(r){
            return aiMap[r.id] ? Object.assign(r, aiMap[r.id]) : r;
          });
        }
      }
    } catch(aiErr) {}
    return results;
  } catch(err) {
    return {error: err.toString()};
  }
}

function testRead() {
  var result = readSheet('Tasks');
  Logger.log(JSON.stringify(result).slice(0, 500));
}

function deleteRow(sheetName, id) {
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return {ok: true};
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return {ok: true};
    var headers = data[0].map(function(h){ return String(h).trim(); });
    var idCol = headers.indexOf('id');
    if (idCol < 0) return {error: 'No id column'};
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][idCol]).trim() === String(id).trim()) {
        sheet.deleteRow(i + 1);
        return {ok: true, deleted: id};
      }
    }
    return {ok: true, notFound: id};
  } catch(err) {
    return {error: err.toString()};
  }
}

function detectLeftCompany(subject, bodyPreview) {
  var text = ((subject || '') + ' ' + (bodyPreview || '')).toLowerCase();
  var patterns = [
    'no longer works here', 'no longer with', 'no longer employed',
    'no longer attend', 'no longer at this address', 'no longer active',
    'is no longer', 'i am no longer', 'i have left', 'has left the company',
    'please update your records', 'please update your systems',
    'this email address is no longer', 'no longer monitored',
    'i have moved on', 'out of the business', 'has left ' ,
    'this mailbox is no longer', 'address is no longer attended'
  ];
  var matched = null;
  for (var i = 0; i < patterns.length; i++) {
    if (text.indexOf(patterns[i]) !== -1) { matched = patterns[i]; break; }
  }
  if (!matched) return { isLeft: false };

  // Try to find a replacement email address near the trigger phrase
  var emailMatch = (bodyPreview || '').match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g);
  var newEmail = null;
  if (emailMatch && emailMatch.length) {
    // Prefer the last email mentioned (usually the redirect address), skip if it matches sender
    newEmail = emailMatch[emailMatch.length - 1];
  }
  return { isLeft: true, reason: matched, newEmail: newEmail };
}

function getEmails() {
  try {
    var props = PropertiesService.getUserProperties();
    var token = props.getProperty('ms_access_token');
    var expiry = props.getProperty('ms_token_expiry');
    if (!token || !expiry || new Date().getTime() > parseInt(expiry)) {
      var refreshToken = props.getProperty('ms_refresh_token');
      if (refreshToken) {
        var refreshed = refreshMsToken(refreshToken);
        if (refreshed.error) return buildAuthResponse();
        token = refreshed.access_token;
      } else {
        return buildAuthResponse();
      }
    }
    var url = 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=50&$orderby=receivedDateTime desc&$select=id,subject,from,receivedDateTime,isRead,bodyPreview,webLink,conversationId';
    var resp = UrlFetchApp.fetch(url, {
      headers: {Authorization: 'Bearer ' + token},
      muteHttpExceptions: true
    });
    var data = JSON.parse(resp.getContentText());
    if (data.error) {
      if (data.error.code === 'InvalidAuthenticationToken') {
        props.deleteProperty('ms_access_token');
        return buildAuthResponse();
      }
      return {error: data.error.message};
    }
    var emails = (data.value || []).map(function(m) {
      var senderEmail = m.from && m.from.emailAddress ? m.from.emailAddress.address : '';
      var bodyPreview = m.bodyPreview || '';
      var leftInfo = detectLeftCompany(m.subject || '', bodyPreview);
      return {
        id: m.id,
        subject: m.subject || '',
        sender: (m.from && m.from.emailAddress ? m.from.emailAddress.name : '') + ' - ' + (m.from && m.from.emailAddress ? m.from.emailAddress.address.split('@')[1] : ''),
        senderEmail: senderEmail,
        received: m.receivedDateTime,
        summary: bodyPreview.slice(0, 200),
        isRead: m.isRead,
        webLink: m.webLink,
        conversationId: m.conversationId || '',
        urgency: leftInfo.isLeft ? 'left' : 'info',
        isProspect: false,
        newContactEmail: leftInfo.newEmail || '',
        leftReason: leftInfo.reason || ''
      };
    });
    var leftOnes = emails.filter(function(e){ return e.urgency === 'left'; });
    if (leftOnes.length) {
      try {
        var rows = leftOnes.map(function(e) {
          return {
            id: 'left_' + e.id,
            senderEmail: e.senderEmail,
            senderName: e.sender,
            subject: e.subject,
            newContactEmail: e.newContactEmail,
            reason: e.leftReason,
            detected: new Date().toISOString()
          };
        });
        upsertSheet('LeftCompany', rows);
      } catch(e2) {}
    }
    return {items: emails};
  } catch(err) {
    return {error: err.toString()};
  }
}

function getSentItems() {
  try {
    var props = PropertiesService.getUserProperties();
    var token = props.getProperty('ms_access_token');
    var expiry = props.getProperty('ms_token_expiry');
    if (!token || !expiry || new Date().getTime() > parseInt(expiry)) {
      var refreshToken = props.getProperty('ms_refresh_token');
      if (refreshToken) {
        var refreshed = refreshMsToken(refreshToken);
        if (refreshed.error) return {error: 'Not authenticated'};
        token = refreshed.access_token;
      } else {
        return {error: 'Not authenticated'};
      }
    }

    var cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - 30);
    var cutoffStr = cutoff.toISOString();

    var sentUrl = 'https://graph.microsoft.com/v1.0/me/mailFolders/sentitems/messages?$top=100&$orderby=sentDateTime desc&$select=id,subject,toRecipients,sentDateTime,bodyPreview,conversationId&$filter=sentDateTime ge ' + cutoffStr;
    var sentResp = UrlFetchApp.fetch(sentUrl, {
      headers: {Authorization: 'Bearer ' + token},
      muteHttpExceptions: true
    });
    var sentData = JSON.parse(sentResp.getContentText());
    if (sentData.error) return {error: sentData.error.message};

    var inboxUrl = 'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=100&$select=conversationId,receivedDateTime&$filter=receivedDateTime ge ' + cutoffStr;
    var inboxResp = UrlFetchApp.fetch(inboxUrl, {
      headers: {Authorization: 'Bearer ' + token},
      muteHttpExceptions: true
    });
    var inboxData = JSON.parse(inboxResp.getContentText());
    var repliedConvIds = {};
    (inboxData.value || []).forEach(function(m) {
      if (m.conversationId) repliedConvIds[m.conversationId] = true;
    });

    var now = new Date();
    var fiveDaysAgo = new Date(now.getTime() - 5 * 24 * 60 * 60 * 1000);

    // First filter to no-reply emails over 5 days
    var filtered = (sentData.value || []).filter(function(m) {
      var sentDate = new Date(m.sentDateTime);
      var hasReply = repliedConvIds[m.conversationId];
      var isOldEnough = sentDate <= fiveDaysAgo;
      var subj = (m.subject || '').toLowerCase();
      var isNoise = subj.indexOf('unsubscribe') >= 0 || subj.indexOf('auto') >= 0;
      return isOldEnough && !hasReply && !isNoise;
    });

    // Dedupe by recipient email, keep most recent sent email per person
    var seenRecipients = {};
    var deduped = [];
    filtered.forEach(function(m) {
      var recipient = m.toRecipients && m.toRecipients[0] ? m.toRecipients[0].emailAddress : {};
      var email = (recipient.address || '').toLowerCase();
      if (!email) return;
      if (!seenRecipients[email]) {
        seenRecipients[email] = true;
        deduped.push(m);
      }
    });

    var noReplies = deduped.map(function(m) {
      var recipient = m.toRecipients && m.toRecipients[0] ? m.toRecipients[0].emailAddress : {};
      var daysSince = Math.floor((now - new Date(m.sentDateTime)) / 86400000);
      return {
        id: m.id,
        subject: m.subject || '',
        recipientName: recipient.name || '',
        recipientEmail: recipient.address || '',
        sentDate: m.sentDateTime,
        daysSince: daysSince,
        summary: (m.bodyPreview || '').slice(0, 200),
        webLink: m.webLink || ''
      };
    });

    return {items: noReplies};
  } catch(err) {
    return {error: err.toString()};
  }
}

function buildAuthResponse() {
  var clientId = '88325ec4-e5cf-467f-a3a0-4b507345483b';
  var tenantId = '8bc70f3f-8de6-4de9-a47a-0a87d9b989f6';
  var redirectUri = ScriptApp.getService().getUrl();
  var authUrl = 'https://login.microsoftonline.com/' + tenantId + '/oauth2/v2.0/authorize' +
    '?client_id=' + clientId +
    '&response_type=code' +
    '&redirect_uri=' + encodeURIComponent(redirectUri) +
    '&scope=' + encodeURIComponent('Mail.Read offline_access') +
    '&response_mode=query';
  return {needsAuth: true, authUrl: authUrl};
}

function refreshMsToken(refreshToken) {
  try {
    var clientId = '88325ec4-e5cf-467f-a3a0-4b507345483b';
    var clientSecret = '[REDACTED_AZURE_CLIENT_SECRET_SEE_LOCAL_COPY]';
    var tenantId = '8bc70f3f-8de6-4de9-a47a-0a87d9b989f6';
    var redirectUri = ScriptApp.getService().getUrl();
    var resp = UrlFetchApp.fetch('https://login.microsoftonline.com/' + tenantId + '/oauth2/v2.0/token', {
      method: 'post',
      payload: {
        grant_type: 'refresh_token',
        client_id: clientId,
        client_secret: clientSecret,
        refresh_token: refreshToken,
        redirect_uri: redirectUri
      },
      muteHttpExceptions: true
    });
    var data = JSON.parse(resp.getContentText());
    if (data.error) return {error: data.error_description};
    var props = PropertiesService.getUserProperties();
    props.setProperty('ms_access_token', data.access_token);
    props.setProperty('ms_token_expiry', String(new Date().getTime() + (data.expires_in - 60) * 1000));
    if (data.refresh_token) props.setProperty('ms_refresh_token', data.refresh_token);
    return {access_token: data.access_token};
  } catch(err) {
    return {error: err.toString()};
  }
}

function handleOAuthCallback(code) {
  try {
    var clientId = '88325ec4-e5cf-467f-a3a0-4b507345483b';
    var clientSecret = '[REDACTED_AZURE_CLIENT_SECRET_SEE_LOCAL_COPY]';
    var tenantId = '8bc70f3f-8de6-4de9-a47a-0a87d9b989f6';
    var redirectUri = ScriptApp.getService().getUrl();
    var resp = UrlFetchApp.fetch('https://login.microsoftonline.com/' + tenantId + '/oauth2/v2.0/token', {
      method: 'post',
      payload: {
        grant_type: 'authorization_code',
        client_id: clientId,
        client_secret: clientSecret,
        code: code,
        redirect_uri: redirectUri
      },
      muteHttpExceptions: true
    });
    var data = JSON.parse(resp.getContentText());
    if (data.error) return 'Auth failed: ' + data.error_description;
    var props = PropertiesService.getUserProperties();
    props.setProperty('ms_access_token', data.access_token);
    props.setProperty('ms_token_expiry', String(new Date().getTime() + (data.expires_in - 60) * 1000));
    if (data.refresh_token) props.setProperty('ms_refresh_token', data.refresh_token);
    return 'success';
  } catch(err) {
    return 'Error: ' + err.toString();
  }
}
function cleanActivitySheet() {
  var ss = SpreadsheetApp.openById('1sPVpASsTB6q8P2kkIPJkO4_qKAL7sk7m6dAnBC4iuYE');
  var sheet = ss.getSheetByName('Activity');
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var dateCol = headers.indexOf('date');
  if (dateCol < 0) return;
  var seen = {};
  var keep = [headers];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var d = String(row[dateCol]).slice(0, 10);
    if (!d || seen[d]) continue;
    seen[d] = true;
    keep.push(row);
  }
  sheet.clearContents();
  sheet.getRange(1, 1, keep.length, keep[0].length).setValues(keep);
  Logger.log('Cleaned. Rows kept: ' + (keep.length - 1));
}