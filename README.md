# Requirement Parsing Add-on

## Overview

This Google Apps Script code provides a Google Docs add-on for parsing requirements and generating a summary. The add-on retrieves requirements data from the Requirement Yogi API, searches for specific keys in the Google Docs document, and generates a summary either in the document or as a sidebar HTML display.

## Getting Started

1. Open your Google Docs document.
2. From the menu, go to "Extensions" -> "Apps Script."
3. Paste the provided code into the script editor.
4. Save the script.

## Usage

### Accessing the Add-on

Click on the "Start" menu in the Google Docs toolbar to access the add-on.


### Displaying Summary

Two functions, `EOF` and `Sidebar`, generate the summary display.

- **Sidebar:** Creates an HTML sidebar with links to requirements.
- **EOF:** Appends a page break and a bullet list of requirements at the end of the document.

### Fetching JSON Data

The data is fetched from the from the Requirement Yogi API, iteratively retrieving results until the total is reached.

## External Dependencies

This add-on relies on the Requirement Yogi API for requirements data. The URL for the API is the following

```
https://ww1.requirementyogi.cloud/nuitdelinfo/search?offset=OFFSET_VAL
```

## Error Handling

- Log any errors using `Logger.log` for debugging purposes.
- Provide a user-friendly message if the API call fails.

## Remaining issues

- The links provided with the Sidetab view can be corrupted because of an encoding problem to HTML

