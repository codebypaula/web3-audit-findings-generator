# Web3 Audit Findings Generator

A Google Apps Script tool designed primarily for importing, formatting, and structuring Web3 audit findings directly into Google Docs.

## 🚀 Main Features

- **Findings Importer:** The core functionality of the script. It allows the auditor to easily paste findings written in Markdown (High, Medium, Low, Gas, QA) and automatically inserts them into the document with proper styling.
- **Syntax Highlighting:** Detects and applies color formatting to inline code and code blocks within the findings. It includes native support for Solidity, JavaScript, TypeScript, Python, and Bash.
- **Scope Organization:** Quickly generates and formats the audited smart contracts scope tables.
- **Custom Sidebar:** Adds a user-friendly "Audit Tools" panel inside the Google Docs interface for a seamless workflow.

## 📁 Project Structure

- `Code.gs`: The backend logic for Google Docs document manipulation, paragraph insertion, and text parsing.
- `Sidebar.html`: The frontend user interface that runs inside the Google Docs sidebar panel.

## ©️ Copyright & Usage Rights

**Created by Paula Perini.**

🛑 **Exclusive Use:** This repository and its source code are for the exclusive use of Paula Perini. Unauthorized copying, distribution, modification, public display, or use of this script, in whole or in part, via any medium, is strictly prohibited. All rights reserved.
