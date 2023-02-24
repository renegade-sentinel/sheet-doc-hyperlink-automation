# AddDocLinksFromSheet

AddDocLinksFromSheet is a Google Apps Script that adds hyperlinks to a Google Doc based on a Google Sheet acting as a database for all preferred internal links in SEO content. The script scans the anchor text and hyperlink columns in the Google Sheet, and looks for the phrases in the anchor text column in the Google Doc column to hyperlink. The script also highlights the paragraph where the link was added, making it easy to spot check, and keeps a log of all internal links added in the "script log" sheet tab.

## Description

This script automates the process of hyperlinking within Google Docs, making it useful when working with large SEO documents. By referencing a Google Sheet database for preferred internal links, the script can quickly and accurately add hyperlinks to the desired text within the document. The script also provides additional functionality, including highlighting the paragraph where the link was added and keeping a log of all internal links added in the "script log" sheet tab.

## Features

- Add hyperlinks to a Google Doc based on a Google Sheet database
- Automatically hyperlink text in bulk
- Streamline the SEO content creation process
- Scans the anchor text and hyperlink columns in the Google Sheet
- Highlights the paragraph where the link was added for easy spot checking
- Logs all internal links added in the "script log" sheet tab

## Getting started

To use AddDocLinksFromSheet, simply clone the repository and run the script. You will need to provide the necessary Google API credentials and authenticate the script to access your Google Drive.

## Contributions

Contributions are welcome! If you'd like to contribute to AddDocLinksFromSheet, please open an issue or a pull request.
