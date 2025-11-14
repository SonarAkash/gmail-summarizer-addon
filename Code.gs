function onGmailMessageOpen(e) {
  const messageId = e.gmail.messageId;
  const message = GmailApp.getMessageById(messageId);
  const emailBody = message.getPlainBody();
  
  if (!emailBody) {
    retufunction onGmailMessageOpen(e) {
  const messageId = e.gmail.messageId;
  const message = GmailApp.getMessageById(messageId);
  const emailBody = message.getPlainBody();

  const globalTime = fetchGlobalTime();

  let summary = '';
  let error = null;
  
  if (!emailBody || emailBody.trim() === '') {
    summary = "This email has no text content to summarize.";
  }else{
    try {
      const prompt = emailBody; 
      summary = callGoogleGemini(prompt);
    } catch (err) {
      Logger.log("Error in onGmailMessageOpen: " + err);
      error = err.message;
    }
  }
  return createSummaryCard(summary, globalTime, error, messageId);
  
}

function onRetrySummarize(e) {
  const messageId = e.parameters.messageId;
  
  if (!messageId) {
    const card = createSummaryCard("Error: Could not find message to retry.", fetchGlobalTime(), null, "0");
    return CardService.newNavigation().updateCard(card);
  }

  const message = GmailApp.getMessageById(messageId);
  const emailBody = message.getPlainBody();
  
  const globalTime = fetchGlobalTime();
  let newCard;
  
  if (!emailBody || emailBody.trim() === '') {
    newCard = createSummaryCard("This email has no text content to summarize.", globalTime, null, messageId);
  } else {
    try {
      const prompt = emailBody; 
      const summary = callGoogleGemini(prompt);
      newCard = createSummaryCard(summary, globalTime, null, messageId); 
      
    } catch (err) {
      Logger.log("Error in onRetrySummarize: " + err);
      newCard = createSummaryCard("", globalTime, err.message, messageId); 
    }
  }
  
  return CardService.newNavigation().updateCard(newCard);
}


function callGoogleGemini(prompt) {
  const API_KEY = PropertiesService.getUserProperties().getProperty('GEMINI_KEY');
  if (!API_KEY) {
    throw new Error("No Api-key");
  }
  
  const GEMINI_URL = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + API_KEY;
  
  const payload = {
    'systemInstruction': {
      'parts': [ 
        { 'text': 'You are a helpful assistant. Summarize the following email in one or two concise sentences.' }
      ]
    },
    'contents': [ 
      { 'parts': [ { 'text': prompt } ] }
    ]
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload), 
    'muteHttpExceptions': true
  };

  const response = UrlFetchApp.fetch(GEMINI_URL, options);
  const responseText = response.getContentText();
  const jsonResponse = JSON.parse(responseText);

  if (jsonResponse.error) {
    Logger.log("Google API Error: " + jsonResponse.error.message);
    throw new Error("Google Error: " + jsonResponse.error.message);
  }

  if (jsonResponse.candidates && jsonResponse.candidates.length > 0 &&
      jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts &&
      jsonResponse.candidates[0].content.parts.length > 0) {
        
    const summary = jsonResponse.candidates[0].content.parts[0].text.trim();
    return summary;
  } else {
    Logger.log("Google Error: " + responseText);
    throw new Error("Google Error: Could not parse summary from API response.");
  }
}



function createSummaryCard(summaryText, globalTimeText, errorText, messageId) {
  const cardBuilder = CardService.newCardBuilder();
  
  if (!errorText) {
    cardBuilder.setHeader(CardService.newCardHeader().setTitle('Email Summary'));
    cardBuilder.addSection(
      CardService.newCardSection()
        .addWidget(
          CardService.newTextParagraph()
            .setText(summaryText)
        )
    );
  }
  
  const replyInput = CardService.newTextInput()
    .setFieldName("reply_content")
    .setTitle("Draft a quick response");

  const draftAction = CardService.newAction()
    .setFunctionName('onDraftReply')
    .setParameters({ messageId: messageId });

  cardBuilder.addSection(
    CardService.newCardSection()
      .setHeader('Respond')
      .addWidget(replyInput)
      .addWidget(
        CardService.newTextButton()
          .setText('Create Draft')
          .setOnClickAction(draftAction)
      )
  );

  cardBuilder.addSection(
    CardService.newCardSection()
      .setHeader('Live Data')
      .addWidget(
        CardService.newTextParagraph()
          .setText(globalTimeText)
      )
  );

  if (errorText) {
    cardBuilder.setHeader(CardService.newCardHeader().setTitle('Error'));
    cardBuilder.addSection(
      CardService.newCardSection()
        .addWidget(
          CardService.newTextParagraph()
            .setText(errorText)
        )
    );

    const retryAction = CardService.newAction()
      .setFunctionName('onRetrySummarize')
      .setParameters({ messageId: messageId });

    cardBuilder.setFixedFooter(
      CardService.newFixedFooter()
        .setPrimaryButton(
          CardService.newTextButton()
            .setText('RETRY')
            .setOnClickAction(retryAction)
        )
    );
  }

  return cardBuilder.build();
}


function onDraftReply(e) {
  const messageId = e.parameters.messageId;
  
  const userReplyText = e.formInput.reply_content;

  if (!userReplyText || userReplyText.trim() === '') {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification()
        .setText('Please type a reply first.'))
      .build();
  }
  
  const message = GmailApp.getMessageById(messageId);
  
  const draft = message.createDraftReply(userReplyText);
  
  return CardService.newComposeActionResponseBuilder()
    .setGmailDraft(draft)
    .build();
}


function storeGeminiApiKey() {
  const GEMINI_API_KEY = 'GEMINI_API_KEY'; 
  
  if (GEMINI_API_KEY === 'GEMINI_API_KEY') {
    Logger.log('No Api-key');
    return;
  }
  
  try {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('GEMINI_KEY', GEMINI_API_KEY);
    Logger.log('SUCCESS: Google Gemini API key has been stored.');
  } catch (e) {
    Logger.log('ERROR: Could not store API key. ' + e);
  }
}

function checkGeminiApiKey() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const storedKey = userProperties.getProperty('GEMINI_KEY');
    
    if (storedKey) {
      Logger.log('SUCCESS: Gemini key is stored. It starts with: ' + storedKey.substring(0, 6));
    } else {
      Logger.log('FAILURE: No Gemini API key is stored.');
    }
  } catch (e) {
    Logger.log('ERROR: Could not retrieve API key. ' + e);
  }
}



function fetchGlobalTime() {
  const TIME_URL = 'http://worldtimeapi.org/api/timezone/Etc/UTC';
  
  try {
    const options = {
      'method': 'get',
      'contentType': 'application/json',
      'muteHttpExceptions': true
    };
    
    const response = UrlFetchApp.fetch(TIME_URL, options);
    const responseText = response.getContentText();
    const jsonResponse = JSON.parse(responseText);
    
    if (jsonResponse.error) {
      throw new Error(jsonResponse.error);
    }
    
    const dateTime = new Date(jsonResponse.datetime);
    const formattedTime = 'Current UTC Time: ' + dateTime.toUTCString();
    
    return formattedTime;
    
  } catch (e) {
    Logger.log('WorldTimeAPI Error: ' + e);
    return 'Could not load global time.';
  }
}rn createSummaryCard("This email has no text content to summarize.", null);
  }

  try {
    const prompt = "Summarize the following email in one or two sentences: \n\n" + emailBody;
    const summary = callGoogleGemini(prompt);
    return createSummaryCard(summary, null); 
    
  } catch (err) {
    Logger.log("Error in onGmailMessageOpen: " + err);
    return createSummaryCard(err.message, messageId);
  }
}

function onRetrySummarize(e) {
  const messageId = e.parameters.messageId; 
  
  if (!messageId) {
    const card = createSummaryCard("Error: Could not find message to retry.", null);
    return CardService.newNavigation().updateCard(card);
  }

  const message = GmailApp.getMessageById(messageId);
  const emailBody = message.getPlainBody();
  
  let newCard;
  try {
    const prompt = "Summarize the following email in one or two sentences: \n\n" + emailBody;
    const summary = callGoogleGemini(prompt);
    newCard = createSummaryCard(summary, null);
    
  } catch (err) {
    Logger.log("Error in onRetrySummarize: " + err);
    newCard = createSummaryCard(err.message, messageId);
  }
  
  return CardService.newNavigation().updateCard(newCard);
}

function callGoogleGemini(prompt) {
  const API_KEY = PropertiesService.getUserProperties().getProperty('GEMINI_KEY');
  if (!API_KEY) {
    throw new Error("No Api-key");
  }
  
  const GEMINI_URL = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + API_KEY;
  
  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify({
      'contents': [
        { 'parts': [ { 'text': prompt } ] }
      ]
    }),
    'muteHttpExceptions': true
  };

  const response = UrlFetchApp.fetch(GEMINI_URL, options);
  const responseText = response.getContentText();
  const jsonResponse = JSON.parse(responseText);

  if (jsonResponse.error) {
    Logger.log("Google API Error: " + jsonResponse.error.message);
    throw new Error("Google Error: " + jsonResponse.error.message);
  }

  if (jsonResponse.candidates && jsonResponse.candidates.length > 0 &&
      jsonResponse.candidates[0].content && jsonResponse.candidates[0].content.parts &&
      jsonResponse.candidates[0].content.parts.length > 0) {
        
    const summary = jsonResponse.candidates[0].content.parts[0].text.trim();
    return summary;
  } else {
    Logger.log("Google Error: " + responseText);
    throw new Error("Google Error: Could not parse summary from API response.");
  }
}

function createSummaryCard(summaryText, messageId) {
  const cardBuilder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle('Email Summary'))
    .addSection(
      CardService.newCardSection()
        .addWidget(
          CardService.newTextParagraph()
            .setText(summaryText)
        )
    );

  if (messageId) {
    const retryAction = CardService.newAction()
      .setFunctionName('onRetrySummarize')
      .setParameters({messageId: messageId});

    cardBuilder.setFixedFooter(
      CardService.newFixedFooter()
        .setPrimaryButton(
          CardService.newTextButton()
            .setText('RETRY')
            .setOnClickAction(retryAction)
        )
    );
  }

  return cardBuilder.build();
}

function storeGeminiApiKey() {
  const GEMINI_API_KEY = 'GEMINI_API_KEY'; 
  
  if (GEMINI_API_KEY === 'GEMINI_API_KEY') {
    Logger.log('No Api-key');
    return;
  }
  
  try {
    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('GEMINI_KEY', GEMINI_API_KEY);
    Logger.log('SUCCESS: Google Gemini API key has been stored.');
  } catch (e) {
    Logger.log('ERROR: Could not store API key. ' + e);
  }
}

function checkGeminiApiKey() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const storedKey = userProperties.getProperty('GEMINI_KEY');
    
    if (storedKey) {
      Logger.log('SUCCESS: Gemini key is stored. It starts with: ' + storedKey.substring(0, 6));
    } else {
      Logger.log('FAILURE: No Gemini API key is stored.');
    }
  } catch (e) {
    Logger.log('ERROR: Could not retrieve API key. ' + e);
  }
}