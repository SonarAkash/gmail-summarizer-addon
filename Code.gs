function onGmailMessageOpen(e) {
  const messageId = e.gmail.messageId;
  const message = GmailApp.getMessageById(messageId);
  const emailBody = message.getPlainBody();
  
  const globalTime = fetchGlobalTimeWithBackup();
  
  let summary = '';
  let error = null;

  if (!emailBody || emailBody.trim() === '') {
    summary = "This email has no text content to summarize.";
  } else {
    try {
      summary = callGoogleGeminiWithBackup(emailBody);
    } catch (err) {
      console.error("All AI Models Failed: " + err);
      summary = "⚠️ AI is currently busy. Please try again."; 
      error = "Server overloaded.";
    }
  }
  
  return createSummaryCard(summary, globalTime, error, messageId);
}

function onRetrySummarize(e) {
  const messageId = e.parameters.messageId;
  
  if (!messageId) {
    const card = createSummaryCard("Error: Could not find message.", fetchGlobalTimeWithBackup(), null, "0");
    return CardService.newNavigation().updateCard(card);
  }

  const message = GmailApp.getMessageById(messageId);
  const emailBody = message.getPlainBody();
  
  const globalTime = fetchGlobalTimeWithBackup();
  let newCard;
  
  if (!emailBody || emailBody.trim() === '') {
    newCard = createSummaryCard("This email has no text content.", globalTime, null, messageId);
  } else {
    try {
      const summary = callGoogleGeminiWithBackup(emailBody);
      newCard = createSummaryCard(summary, globalTime, null, messageId); 
    } catch (err) {
      newCard = createSummaryCard(
        "⚠️ Summary unavailable right now.", 
        globalTime, 
        "Still busy. Please wait 5s and retry.", 
        messageId
      ); 
    }
  }
  
  return CardService.newNavigation().updateCard(newCard);
}

function onDraftReply(e) {
  const messageId = e.parameters.messageId;
  const userReplyText = e.formInput.reply_content;

  if (!userReplyText || userReplyText.trim() === '') {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText('Please type a reply first.'))
      .build();
  }
  
  const message = GmailApp.getMessageById(messageId);
  const draft = message.createDraftReply(userReplyText);
  
  return CardService.newComposeActionResponseBuilder()
    .setGmailDraft(draft)
    .build();
}


function callGoogleGeminiWithBackup(prompt) {
  const API_KEY = PropertiesService.getUserProperties().getProperty('GEMINI_KEY');
  if (!API_KEY) throw new Error("No Api-key");

  const MODELS = [
    'gemini-2.5-flash', // Priority 1: Fastest
    'gemini-1.5-flash', // Priority 2: Stable
    'gemini-1.0-pro'    // Priority 3: Legacy
  ];

  let lastError = null;

  for (const model of MODELS) {
    try {
      Utilities.sleep(500);
      return makeGeminiRequest(model, prompt, API_KEY);
    } catch (e) {
      console.warn(`Model ${model} failed, switching to next...`);
      lastError = e;
    }
  }
  
  throw new Error("All AI backups failed.");
}

function makeGeminiRequest(model, prompt, key) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${key}`;
  
  const payload = {
    'systemInstruction': {
      'parts': [{ 'text': 'You are a helpful assistant. Summarize the following email in one or two concise sentences.' }]
    },
    'contents': [{ 'parts': [ { 'text': prompt } ] }]
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (json.error) {
    throw new Error(json.error.message);
  }
  
  if (json.candidates && json.candidates[0].content) {
    return json.candidates[0].content.parts[0].text.trim();
  }
  throw new Error("Empty response");
}

function fetchGlobalTimeWithBackup() {
  try {
    const url = 'http://worldtimeapi.org/api/timezone/Etc/UTC';
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      const date = new Date(json.datetime);
      return "UTC (API): " + date.toUTCString().replace('GMT', '');
    }
  } catch (e) {
    console.warn("Time API failed, using system backup.");
  }

  const now = new Date();
  return "UTC (System): " + now.toUTCString().replace('GMT', '');
}


function createSummaryCard(summaryText, globalTimeText, errorText, messageId) {
  const card = CardService.newCardBuilder();
  
  const mainHeader = CardService.newCardHeader()
    .setTitle('Gmail Summarizer')
    .setImageStyle(CardService.ImageStyle.CIRCLE)
    .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/auto_awesome_black_24dp.png");
  
  card.setHeader(mainHeader);

  if (!errorText) {
    card.addSection(CardService.newCardSection()
      .setHeader("AI Summary") 
      .addWidget(CardService.newTextParagraph().setText(summaryText))
    );
  }
  
  const replyInput = CardService.newTextInput()
    .setFieldName("reply_content")
    .setTitle("Draft a quick response");

  const draftAction = CardService.newAction().setFunctionName('onDraftReply').setParameters({ messageId: messageId });
  
  card.addSection(CardService.newCardSection()
    .setHeader("Respond")
    .addWidget(replyInput)
    .addWidget(CardService.newTextButton()
      .setText('Open Draft')
      .setIcon(CardService.Icon.EMAIL)
      .setOnClickAction(draftAction)
    )
  );

  card.addSection(CardService.newCardSection()
    .addWidget(
      CardService.newDecoratedText()
        .setText(globalTimeText)
        .setTopLabel("Live Global Time")
        .setStartIcon(
          CardService.newIconImage()
            .setIconUrl("https://www.gstatic.com/images/icons/material/system_gm/1x/schedule_black_24dp.png")
        )
    )
  );

  if (errorText) {
    const retryAction = CardService.newAction().setFunctionName('onRetrySummarize').setParameters({ messageId: messageId });
    
    card.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph().setText(errorText))
      .addWidget(CardService.newTextButton()
        .setText('Try Again')
        .setOnClickAction(retryAction)
      )
    );
  }

  return card.build();
}


function storeGeminiApiKey() {
  const GEMINI_API_KEY = 'GEMINI_API_KEY'; 
  if (GEMINI_API_KEY === 'GEMINI_API_KEY') { Logger.log('No Api-key'); return; }
  PropertiesService.getUserProperties().setProperty('GEMINI_KEY', GEMINI_API_KEY);
  Logger.log('Key stored.');
}

function checkGeminiApiKey() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const storedKey = userProperties.getProperty('GEMINI_KEY');
    if (storedKey) { Logger.log('SUCCESS: Key stored.'); } 
    else { Logger.log('FAILURE: No Key.'); }
  } catch (e) { Logger.log('ERROR: ' + e); }
}