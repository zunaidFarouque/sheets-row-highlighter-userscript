# Sheets Row Highlighter (Userscript Port)

A lightweight Tampermonkey userscript to highlight the row and column of the current cell in Google Sheets and Microsoft Excel Online.

This is a highly optimized, zero-dependency port of the original [Sheets Row Highlighter extension](https://github.com/matsu7089/sheets-row-highlighter) by matsu7089. It strips out the popup UI and extension storage overhead, relying entirely on native DOM manipulation and Tampermonkey's built-in script toggling for maximum performance.

## Installation

You can install this script directly via GreasyFork:
**[Install from GreasyFork](insert_your_greasyfork_link_here)**

### Requirements
* A userscript manager like [Tampermonkey](https://www.tampermonkey.net/) installed in your browser.

## Configuration

If you want to change the highlight color, opacity, or toggle row/column functionality, simply open the script in your Tampermonkey dashboard and edit the `CONFIG` block at the top of the file:

```javascript
const CONFIG = {
    color: '#0e65eb', // Hex color code
    opacity: '0.1',   // 0.01 to 0.5 recommended
    enableRow: true,  // Set to true or false
    enableCol: false  // Set to true or false
};
```

## Credits & License

* Core logic and original Chrome/Firefox extension created by [matsu7089](https://github.com/matsu7089).
* Licensed under the [MIT License](https://www.google.com/search?q=LICENSE).
