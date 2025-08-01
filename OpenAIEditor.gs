/**
 * Retrieve the OpenAI API key from script properties. This prevents the key
 * from being stored directly in the source code where it could be exposed.
 */
function getApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) throw new Error('OPENAI_API_KEY is not configured');
  return key;
}

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

  // Read the current document text and styles
  const { text: originalText, images, styles } = docToText(body);

  const prompt =
    `Apply the following instructions to the text and return updated text only.` +
    `\n\nInstructions:\n${instructions}\n\nDocument Text:\n${originalText}`;
  const newText = callChatGPT(prompt);

  // Apply the returned text back into the document, preserving style and images
  applyText(body, newText, images, styles);
  logChange(originalText, newText);
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
      Authorization: `Bearer ${getApiKey()}`
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

// Read the document body text and capture images and basic styles
function docToText(body) {
  const lines = [];
  const images = [];
  const styles = [];
  for (let i = 0; i < body.getNumChildren(); i++) {
    const elem = body.getChild(i);
    const type = elem.getType();
    if (type === DocumentApp.ElementType.PARAGRAPH) {
      const p = elem.asParagraph();
      if (
        p.getNumChildren() === 1 &&
        (p.getChild(0).getType() === DocumentApp.ElementType.INLINE_IMAGE ||
          p.getChild(0).getType() === DocumentApp.ElementType.INLINE_DRAWING)
      ) {
        const placeholder = `[[IMAGE_${images.length}]]`;
        images.push(p.getChild(0).getBlob());
        lines.push(placeholder);
        styles.push(null);
      } else {
        const te = p.editAsText();
        const text = te.getText();
        const style = text.length > 0 ? te.getTextStyle(0, text.length - 1) : null;
        lines.push(text);
        styles.push({ textStyle: style, heading: p.getHeading() });
      }
    } else if (
      type === DocumentApp.ElementType.INLINE_IMAGE ||
      type === DocumentApp.ElementType.INLINE_DRAWING
    ) {
      const placeholder = `[[IMAGE_${images.length}]]`;
      images.push(elem.getBlob());
      lines.push(placeholder);
      styles.push(null);
    }
  }
  return { text: lines.join('\n'), images, styles };
}

// Apply plain text back into the document body preserving style information
function applyText(body, text, images, styles) {
  body.clear();
  const lines = text.split(/\r?\n/);
  let styleIndex = 0;
  lines.forEach((line) => {
    const imgMatch = line.trim().match(/^\[\[IMAGE_(\d+)\]\]$/);
    if (imgMatch) {
      const idx = Number(imgMatch[1]);
      if (images[idx]) body.appendImage(images[idx]);
    } else {
      const style = styles[styleIndex++] || {};
      const p = body.appendParagraph('');
      if (style.heading) p.setHeading(style.heading);
      const te = p.editAsText();
      te.setText(line);
      if (style.textStyle && line.length > 0) {
        te.setTextStyle(0, line.length - 1, style.textStyle);
      }
    }
  });
}
