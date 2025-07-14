# openai-googleapps

This repository contains an example Google Apps Script for integrating the OpenAI API with Google Docs.
The script no longer exports your document to Markdown. Instead it reads the document structure directly using `DocumentApp` and sends the plain text to OpenAI. When the response is received, the script updates the document using `Text.setText()` and `Text.setTextStyle()` so the original formatting is preserved. Embedded images and drawings are inserted back into their original positions.

## Files

- `OpenAIEditor.gs` – core Google Apps Script code.
- `Sidebar.html` – HTML for the sidebar UI.

## Usage

1. Copy the contents of `OpenAIEditor.gs` and `Sidebar.html` into a new Google Apps Script project bound to a Google Doc.
2. Replace `YOUR_API_KEY_HERE` in `OpenAIEditor.gs` with your OpenAI API key.
3. Reload the document. A new **OpenAI Tools** menu will appear allowing you to open the assistant sidebar.
4. Use the sidebar to enter instructions. The script reads the document text directly, sends it along with your instructions to OpenAI, and then applies the returned text back into the document while restoring the original formatting.
5. Changes are logged in the document properties under `change_logs` for simple version tracking.
