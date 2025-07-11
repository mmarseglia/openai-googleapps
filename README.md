# openai-googleapps

This repository contains an example Google Apps Script for integrating the OpenAI API with Google Docs.
The script now converts your document to **Markdown** before sending it to OpenAI and applies the returned Markdown back into the document. This helps preserve basic formatting such as headings, bold and italic text.

## Files

- `OpenAIEditor.gs` – core Google Apps Script code.
- `Sidebar.html` – HTML for the sidebar UI.

## Usage

1. Copy the contents of `OpenAIEditor.gs` and `Sidebar.html` into a new Google Apps Script project bound to a Google Doc.
2. Replace `YOUR_API_KEY_HERE` in `OpenAIEditor.gs` with your OpenAI API key.
3. Reload the document. A new **OpenAI Tools** menu will appear allowing you to open the assistant sidebar.
4. Use the sidebar to enter instructions. The script converts the current document into Markdown, sends it along with your instructions to OpenAI and then applies the returned Markdown back into the document.
5. Changes are logged in the document properties under `change_logs` for simple version tracking.
