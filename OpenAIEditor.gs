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

  // Convert the existing Google Doc to Markdown so formatting is preserved
  const originalMarkdown = docToMarkdown(body);

  const prompt =
    `Apply the following instructions to the Markdown text and return updated Markdown only.` +
    `\n\nInstructions:\n${instructions}\n\nDocument Markdown:\n${originalMarkdown}`;
  const newMarkdown = callChatGPT(prompt);

  // Apply the returned Markdown back into the document
  applyMarkdown(body, newMarkdown);
  logChange(originalMarkdown, newMarkdown);
}

function callChatGPT(prompt) {
  const url = 'https://api.openai.com/v1/chat/completions';

  const payload = {
    model: 'gpt-4.1',
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

// Convert the document body to Markdown, capturing basic formatting
function docToMarkdown(body) {
  const paragraphs = body.getParagraphs();
  return paragraphs.map(paragraphToMarkdown).join('\n\n');
}

function paragraphToMarkdown(p) {
  let text = textElementToMarkdown(p.editAsText());
  switch (p.getHeading()) {
    case DocumentApp.ParagraphHeading.HEADING1:
      return '# ' + text;
    case DocumentApp.ParagraphHeading.HEADING2:
      return '## ' + text;
    case DocumentApp.ParagraphHeading.HEADING3:
      return '### ' + text;
    default:
      return text;
  }
}

function textElementToMarkdown(te) {
  const txt = te.getText();
  let md = '';
  let prevBold = false;
  let prevItalic = false;
  for (let i = 0; i < txt.length; i++) {
    const curBold = te.isBold(i);
    const curItalic = te.isItalic(i);
    if (curBold !== prevBold) {
      md += '**';
      prevBold = curBold;
    }
    if (curItalic !== prevItalic) {
      md += '_';
      prevItalic = curItalic;
    }
    md += txt[i];
  }
  if (prevBold) md += '**';
  if (prevItalic) md += '_';
  return md;
}

// Apply Markdown text back into the document body
function applyMarkdown(body, markdown) {
  body.clear();
  const lines = markdown.split(/\r?\n/);
  lines.forEach(line => {
    if (line.startsWith('# ')) {
      body.appendParagraph(line.substring(2)).setHeading(DocumentApp.ParagraphHeading.HEADING1);
    } else if (line.startsWith('## ')) {
      body.appendParagraph(line.substring(3)).setHeading(DocumentApp.ParagraphHeading.HEADING2);
    } else if (line.startsWith('### ')) {
      body.appendParagraph(line.substring(4)).setHeading(DocumentApp.ParagraphHeading.HEADING3);
    } else if (line.trim() === '') {
      body.appendParagraph('');
    } else {
      appendInlineMarkdown(body.appendParagraph(''), line);
    }
  });
}

function appendInlineMarkdown(paragraph, text) {
  const te = paragraph.editAsText();
  const tokens = parseInlineMarkdown(text);
  tokens.forEach(token => {
    const start = te.getText().length;
    te.insertText(start, token.text);
    const end = te.getText().length - 1;
    if (token.bold) te.setBold(start, end, true);
    if (token.italic) te.setItalic(start, end, true);
  });
}

function parseInlineMarkdown(text) {
  const tokens = [];
  let i = 0;
  let bold = false;
  let italic = false;
  let chunk = '';
  while (i < text.length) {
    if (text.startsWith('**', i)) {
      if (chunk) {
        tokens.push({ text: chunk, bold, italic });
        chunk = '';
      }
      bold = !bold;
      i += 2;
    } else if (text[i] === '_') {
      if (chunk) {
        tokens.push({ text: chunk, bold, italic });
        chunk = '';
      }
      italic = !italic;
      i += 1;
    } else {
      chunk += text[i];
      i += 1;
    }
  }
  if (chunk) tokens.push({ text: chunk, bold, italic });
  return tokens;
}
