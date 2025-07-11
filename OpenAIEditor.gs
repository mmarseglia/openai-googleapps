const OPENAI_API_KEY = 'YOUR_API_KEY_HERE';

function onOpen() {
  DocumentApp.getUi()
    .createMenu('OpenAI Tools')
    .addItem('Open Assistant', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('OpenAI Assistant');
  DocumentApp.getUi().showSidebar(html);
}

function processPrompt(instructions) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const originalText = body.getText();

  const prompt = `Apply the following instructions to the text and return the full updated document.\n\nInstructions:\n${instructions}\n\nDocument text:\n${originalText}`;
  const newText = callChatGPT(prompt);

  body.setText(newText);
  logChange(originalText, newText);
}

function callChatGPT(prompt) {
  const url = 'https://api.openai.com/v1/chat/completions';

  const payload = {
    model: 'gpt-4',
    messages: [{ role: 'user', content: prompt }],
    temperature: 0.7
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${OPENAI_API_KEY}`
    },
    payload: JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());
  return json.choices[0].message.content;
}

function logChange(oldText, newText) {
  const props = PropertiesService.getDocumentProperties();
  const existing = props.getProperty('change_logs') || '';
  const entry = [
    'Time: ' + new Date().toISOString(),
    '--- Old Text ---',
    oldText,
    '--- New Text ---',
    newText,
    ''
  ].join('\n');
  props.setProperty('change_logs', existing + entry + '\n');
}
