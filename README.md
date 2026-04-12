# 🛡️ Web3 Audit Findings Generator

> An automation tool designed to streamline the creation of security reports (Smart Contracts/Web3) directly in Google Docs.

## 🎯 The Problem
In the Smart Contract auditing space, auditors waste precious hours formatting reports manually. Copying and pasting findings from a Markdown file to Google Docs—ensuring formatting, severity levels, and mandatory sections are correct—is a time-consuming and error-prone process.

## 💡 The Solution
The **Web3 Audit Findings Generator** is a native integration built with Google Apps Script that acts as a smart bridge between Markdown notes and the official audit document. 

With a single click, the script parses the text structure, organizes findings by severity, and injects them into Google Docs, strictly maintaining the required visual standards.

## ✨ Key Features

- **Smart Markdown Parsing:** Converts Markdown structures (High, Medium, Low, etc.) into formatted JSON objects.
- **Dynamic Cleanup (Smart Null-handling):** The code automatically detects if the auditor filled in "N/A" or left sections blank (like *Proof of Concept* or *Impact*) and omits them from the final report to maintain document flow.
- **Template Protection:** Validates data before writing to the page, alerting the user via the UI about any missed fields.
- **Visual Formatter ("Fake Heading"):** Applies exact text styles (bolding, font sizes) to simulate subheadings without cluttering the Google Docs Document Outline (navigation bar).

## 🛠️ Technologies Used

- **JavaScript (ES6+)** for parsing logic and text manipulation via Regular Expressions (Regex).
- **Google Apps Script API** for direct UI integration and Google Docs DOM manipulation.
- **HTML / CSS** for building the user interface (Sidebar).
- **Clasp** for version control and deployment from a local development environment (VS Code).

## 🚀 How to Use (For Developers)

If you want to clone this repository and use it in your own Google Docs:

1. Install Google Clasp globally:
   npm install -g @google/clasp

2. Log in to your Google account:
   clasp login

3. Create a new Apps Script project linked to a Google Doc:
   clasp create --type docs --title "Audit Generator"

4. Push the local code to the cloud:
   clasp push

## 👩‍💻 Author

**Paula (codebypaula)** *Project Manager & Developer* Passionate about optimizing processes through code in the Blockchain ecosystem.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
