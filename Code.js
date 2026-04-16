/* =========================================================
   GOOGLE DOCS MENU / SIDEBAR
   ========================================================= */

// Adds a custom menu when the document is opened
function onOpen() {
  DocumentApp.getUi()
    .createMenu('Audit Tools')
    .addItem('Open Import Panel', 'showSidebar')
    .addToUi();
}

// Opens the import sidebar in Google Docs
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Import Findings');
  DocumentApp.getUi().showSidebar(html);
}

/* =========================================================
   PARSING - STRUCTURED INPUT
   ========================================================= */

// Extracts a single-line field like "severity: High"
function extractField(text, fieldName) {
  const regex = new RegExp(fieldName + '\\s*:\\s*(.+)', 'i');
  const match = text.match(regex);
  return match ? match[1].trim() : '';
}

// Extracts a full section until the next known section starts
function extractSection(text, sectionName, nextSections) {
  const nextPattern = nextSections.map(s => s + '\\s*:').join('|');
  const regex = new RegExp(
    sectionName + '\\s*:\\s*([\\s\\S]*?)(?=' + nextPattern + '|$)',
    'i'
  );
  const match = text.match(regex);
  return match ? match[1].trim() : '';
}

// Parses findings from the structured input format
function parseFindings(text) {
  const blocks = text.split('=== FINDING START ===')
    .map(b => b.trim())
    .filter(b => b && b.includes('=== FINDING END ==='));

  const findings = [];

  for (const block of blocks) {
    const clean = block.replace('=== FINDING END ===', '').trim();

    findings.push({
      severity: extractField(clean, 'severity'),
      id: extractField(clean, 'id'),
      title: extractField(clean, 'title'),
      description: extractSection(clean, 'description', [
        'vulnerability_details',
        'proof_of_concept',
        'impact',
        'recommendation',
        'status'
      ]),
      vulnerability_details: extractSection(clean, 'vulnerability_details', [
        'proof_of_concept',
        'impact',
        'recommendation',
        'status'
      ]),
      proof_of_concept: extractSection(clean, 'proof_of_concept', [
        'impact',
        'recommendation',
        'status'
      ]),
      impact: extractSection(clean, 'impact', [
        'recommendation',
        'status'
      ]),
      recommendation: extractSection(clean, 'recommendation', [
        'status'
      ]),
      status: extractField(clean, 'status')
    });
  }

  return findings;
}

/* =========================================================
   PARSING - MARKDOWN INPUT
   ========================================================= */

// Normalizes markdown severity names to the expected format
function normalizeMarkdownSeverity(name) {
  const cleaned = String(name || '').trim().toLowerCase();

  if (cleaned === 'high') return 'High';
  if (cleaned === 'medium') return 'Medium';
  if (cleaned === 'low') return 'Low';
  if (cleaned === 'gas') return 'Gas';
  if (cleaned === 'qa') return 'QA';

  return '';
}

// Normalizes markdown section titles to internal field names
function normalizeMarkdownSectionName(name) {
  const cleaned = String(name || '').trim().toLowerCase();

  if (cleaned === 'description') return 'description';
  if (cleaned === 'vulnerability details') return 'vulnerability_details';
  if (cleaned === 'proof of concept') return 'proof_of_concept';
  if (cleaned === 'impact') return 'impact';
  if (cleaned === 'recommendation') return 'recommendation';
  if (cleaned === 'status') return 'status';

  return '';
}

// Parses findings from markdown headings and sections
function parseMarkdownFindings(markdown) {
  const lines = String(markdown || '').replace(/\r\n/g, '\n').split('\n');

  const findings = [];
  let currentSeverity = '';
  let currentFinding = null;
  let currentSection = '';

  // Pushes the current finding into the result list
  function pushCurrentFinding() {
    if (currentFinding) {
      Object.keys(currentFinding).forEach(key => {
        if (typeof currentFinding[key] === 'string') {
          currentFinding[key] = currentFinding[key].trim();
        }
      });
      findings.push(currentFinding);
      currentFinding = null;
    }
  }

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    const severityMatch = line.match(/^##\s+(High|Medium|Low|Gas|QA)\s*$/i);
    if (severityMatch) {
      pushCurrentFinding();
      currentSeverity = normalizeMarkdownSeverity(severityMatch[1]);
      currentSection = '';
      continue;
    }

    const findingMatch = line.match(/^###\s+\[(.+?)\]\s+(.+?)\s*$/);
    if (findingMatch) {
      pushCurrentFinding();

      currentFinding = {
        severity: currentSeverity,
        id: findingMatch[1].trim(),
        title: findingMatch[2].trim(),
        description: '',
        vulnerability_details: '',
        proof_of_concept: '',
        impact: '',
        recommendation: '',
        status: ''
      };

      currentSection = '';
      continue;
    }

    const sectionMatch = line.match(/^####\s+(.+?)\s*$/);
    if (sectionMatch && currentFinding) {
      currentSection = normalizeMarkdownSectionName(sectionMatch[1]);
      continue;
    }

    if (currentFinding && currentSection) {
      currentFinding[currentSection] +=
        (currentFinding[currentSection] ? '\n' : '') + line;
    }
  }

  pushCurrentFinding();
  return findings;
}

// Converts parsed findings back into the structured text format
function findingsToStructuredText(findings) {
  return findings.map(f => {
    return [
      '=== FINDING START ===',
      `severity: ${f.severity}`,
      `id: ${f.id}`,
      `title: ${f.title}`,
      'description:',
      f.description || '',
      '',
      'vulnerability_details:',
      f.vulnerability_details || '',
      '',
      'proof_of_concept:',
      f.proof_of_concept || '',
      '',
      'impact:',
      f.impact || '',
      '',
      'recommendation:',
      f.recommendation || '',
      '',
      `status: ${f.status}`,
      '=== FINDING END ==='
    ].join('\n');
  }).join('\n\n');
}

/* =========================================================
   VALIDATION
   ========================================================= */

// Validates severity, status, title, and ID prefix rules
function validateFindings(findings) {
  const allowedSeverities = ['High', 'Medium', 'Low', 'Gas', 'QA'];
  const allowedStatus = ['Resolved', 'Acknowledged', 'Unresolved'];

  const prefix = {
    High: 'H-',
    Medium: 'M-',
    Low: 'L-',
    Gas: 'G-',
    QA: 'Q-'
  };

  const errors = [];

  findings.forEach((f, i) => {
    const n = i + 1;

    if (!allowedSeverities.includes(f.severity)) {
      errors.push(`Finding ${n}: invalid severity (${f.severity})`);
    }

    if (!f.title) {
      errors.push(`Finding ${n}: missing title`);
    }

    if (f.severity && f.id && prefix[f.severity] && !f.id.startsWith(prefix[f.severity])) {
      errors.push(`Finding ${n}: id ${f.id} does not match severity ${f.severity}`);
    }

    if (f.status && !allowedStatus.includes(f.status)) {
      errors.push(`Finding ${n}: invalid status (${f.status})`);
    }
  });

  return errors;
}

// Builds warnings for empty sections inside findings
function buildSectionWarnings(findings) {
  const warnings = [];

  findings.forEach((f, i) => {
    const label = f.id ? f.id : `Finding ${i + 1}`;

    if (isBlank(f.description)) {
      warnings.push(`${label}: Description is empty`);
    }

    if (isBlank(f.vulnerability_details)) {
      warnings.push(`${label}: Vulnerability Details is empty`);
    }

    if (isBlank(f.proof_of_concept)) {
      warnings.push(`${label}: Proof of Concept is empty`);
    }

    if (isBlank(f.impact)) {
      warnings.push(`${label}: Impact is empty`);
    }

    if (isBlank(f.recommendation)) {
      warnings.push(`${label}: Recommendation is empty`);
    }

    if (isBlank(f.status)) {
      warnings.push(`${label}: Status is empty`);
    }
  });

  return warnings;
}

// Validates structured input text and returns a summary message
function validateFindingsText(text) {
  try {
    const findings = parseFindings(text);
    if (!findings.length) return 'No findings found.';

    const errors = validateFindings(findings);
    if (errors.length) return 'Errors:\n\n' + errors.join('\n');

    const counts = countFindingsBySeverity(findings);
    const warnings = buildSectionWarnings(findings);

    const lines = [
      'Structured input OK.',
      '',
      `Total findings: ${findings.length}`,
      `High: ${counts.High}`,
      `Medium: ${counts.Medium}`,
      `Low: ${counts.Low}`,
      `Gas: ${counts.Gas}`,
      `QA: ${counts.QA}`
    ];

    if (warnings.length) {
      lines.push('');
      lines.push('Warnings:');
      lines.push('');
      warnings.forEach(w => lines.push(`- ${w}`));
    }

    return lines.join('\n');
  } catch (e) {
    return 'Error: ' + e.message;
  }
}

// Validates markdown input and returns a summary message
function validateMarkdownText(markdown) {
  try {
    const findings = parseMarkdownFindings(markdown);

    if (!findings.length) {
      return 'No findings found in markdown.';
    }

    const errors = validateFindings(findings);
    if (errors.length) {
      return 'Markdown validation errors:\n\n' + errors.join('\n');
    }

    const counts = countFindingsBySeverity(findings);
    const warnings = buildSectionWarnings(findings);

    const lines = [
      'Markdown OK.',
      '',
      `Total findings: ${findings.length}`,
      `High: ${counts.High}`,
      `Medium: ${counts.Medium}`,
      `Low: ${counts.Low}`,
      `Gas: ${counts.Gas}`,
      `QA: ${counts.QA}`
    ];

    if (warnings.length) {
      lines.push('');
      lines.push('Warnings:');
      lines.push('');
      warnings.forEach(w => lines.push(`- ${w}`));
    }

    return lines.join('\n');
  } catch (e) {
    return 'Markdown validation error: ' + e.message;
  }
}

/* =========================================================
   FIND HELPERS
   ========================================================= */

// Finds the exact paragraph index of a marker
function findMarkerIndex(body, markerText) {
  const total = body.getNumChildren();

  for (let i = 0; i < total; i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const txt = el.asParagraph().getText().trim();
    if (txt === markerText) return i;
  }

  return -1;
}

// Finds the first paragraph that contains the given text
function findParagraphIndexContaining(body, searchText) {
  const total = body.getNumChildren();

  for (let i = 0; i < total; i++) {
    const el = body.getChild(i);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;

    const txt = el.asParagraph().getText().trim();
    if (txt.includes(searchText)) return i;
  }

  return -1;
}

// Checks if a value is empty or only whitespace
function isBlank(value) {
  return !value || !String(value).trim();
}

/* =========================================================
   HTML / CODE HELPERS
   ========================================================= */

// Decodes common HTML entities into readable text
function decodeHtmlEntities(str) {
  if (!str) return '';

  return String(str)
    .replace(/&#x([0-9a-f]+);/gi, function (_, hex) {
      return String.fromCharCode(parseInt(hex, 16));
    })
    .replace(/&#(\d+);/g, function (_, dec) {
      return String.fromCharCode(parseInt(dec, 10));
    })
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/gi, "'")
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&');
}

// Converts RGB values to a hex color string
function rgbToHex(r, g, b) {
  const toHex = (n) => ('0' + parseInt(n, 10).toString(16)).slice(-2);
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

// Parses highlighted HTML code into plain text and color ranges
function parseHtmlCodeToSegments(html) {
  var source = html || '';

  source = source
    .replace(/\r\n/g, '\n')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<p[^>]*>/gi, '')
    .replace(/<pre[^>]*>/gi, '')
    .replace(/<\/pre>/gi, '')
    .replace(/<code[^>]*>/gi, '')
    .replace(/<\/code>/gi, '');

  var tokenRegex = /(<span[^>]*>|<\/span>|<[^>]+>|[^<]+)/gi;

  var match;
  var plainText = '';
  var colorStack = [];
  var segments = [];

  while ((match = tokenRegex.exec(source)) !== null) {
    var token = match[0];

    // Starts a colored span block
    if (/^<span/i.test(token)) {
      var colorMatch = token.match(/color:\s*rgb\((\d+),\s*(\d+),\s*(\d+)\)/i);
      if (colorMatch) {
        var hex = rgbToHex(colorMatch[1], colorMatch[2], colorMatch[3]);
        colorStack.push({ color: hex, start: plainText.length });
      } else {
        // Keeps the span stack balanced even if no color is present
        colorStack.push(null);
      }
      continue;
    }

    // Closes the current colored span
    if (/^<\/span>/i.test(token)) {
      var top = colorStack.pop();
      if (top && top.color !== null && plainText.length > top.start) {
        segments.push({
          start: top.start,
          end: plainText.length - 1,
          color: top.color
        });
      }
      continue;
    }

    // Ignores all other HTML tags
    if (/^</.test(token)) continue;

    // Appends regular text content
    var textChunk = decodeHtmlEntities(token);
    if (textChunk) {
      plainText += textChunk;
    }
  }

  return { plainText: plainText, segments: segments };
}

// Applies the standard style used for code tables
function styleCodeTable(table) {
  const cell = table.getCell(0, 0);
  const text = cell.editAsText();

  text.setFontFamily('Roboto Mono');
  text.setFontSize(9);

  cell.setBackgroundColor('#f5f5f5');
  cell.setPaddingTop(10);
  cell.setPaddingBottom(10);
  cell.setPaddingLeft(10);
  cell.setPaddingRight(10);

  table.setBorderWidth(1);
  table.setBorderColor('#bdbdbd');

  // Applies line spacing to paragraphs inside the code cell
  const numParagraphs = cell.getNumChildren();
  for (let i = 0; i < numParagraphs; i++) {
    const child = cell.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      child.asParagraph().setLineSpacing(1.5);
    }
  }

  return { cell, text };
}

// Inserts a plain code block as a one-cell table
function insertPlainCodeBlockAfter(body, anchorIndex, code) {
  const table = body.insertTable(anchorIndex + 1, [[code || '']]);
  styleCodeTable(table);
  return anchorIndex + 1;
}

// Inserts a highlighted code block and preserves token colors
function insertHighlightedCodeBlockAfter(body, anchorIndex, htmlCode) {
  const parsed = parseHtmlCodeToSegments(htmlCode);

  const table = body.insertTable(anchorIndex + 1, [[parsed.plainText || '']]);
  const styled = styleCodeTable(table);
  const text = styled.text;

  parsed.segments.forEach(seg => {
    if (seg.start >= 0 && seg.end >= seg.start) {
      text.setForegroundColor(seg.start, seg.end, seg.color);
    }
  });

  return anchorIndex + 1;
}

// Automatically decides whether the code block is plain or highlighted HTML
function insertCodeBlockAuto(body, anchorIndex, codeOrHtml) {
  const isHighlightedHtml =
    /<span[^>]*style="[^"]*color:\s*rgb\(/i.test(codeOrHtml || '') ||
    /<code[^>]*>/i.test(codeOrHtml || '') ||
    /<pre[^>]*>/i.test(codeOrHtml || '');

  if (isHighlightedHtml) {
    return insertHighlightedCodeBlockAfter(body, anchorIndex, codeOrHtml);
  }

  return insertPlainCodeBlockAfter(body, anchorIndex, codeOrHtml);
}

/* =========================================================
   MIXED CONTENT HANDLING
   ========================================================= */

// Splits content into text blocks and code blocks
function parseMixedContent(content) {
  var blocks = [];
  var remaining = content;
  var result = [];

  // Splits HTML code blocks first
  var htmlCodeRegex = /(<pre[^>]*><code[^>]*>[\s\S]*?<\/code><\/pre>)/;
  var parts = remaining.split(htmlCodeRegex);

  for (var i = 0; i < parts.length; i++) {
    var part = parts[i];

    if (!part) continue;

    // Captured HTML code blocks
    if (i % 2 === 1) {
      blocks.push({ type: 'code', language: '', content: part });
      continue;
    }

    // Regular text that may still contain fenced code blocks
    var fenceRegex = /```([a-zA-Z0-9_-]+)?\n([\s\S]*?)```/g;
    var lastIndex = 0;
    var match;

    while ((match = fenceRegex.exec(part)) !== null) {
      var textBefore = part.slice(lastIndex, match.index).trim();
      if (textBefore) {
        blocks.push({ type: 'text', content: textBefore });
      }
      blocks.push({ type: 'code', language: (match[1] || '').trim(), content: match[2].trim() });
      lastIndex = fenceRegex.lastIndex;
    }

    var tail = part.slice(lastIndex).trim();
    if (tail) {
      blocks.push({ type: 'text', content: tail });
    }
  }

  return blocks;
}

// Inserts a regular styled text paragraph
function insertTextParagraphAfter(body, anchorIndex, textContent) {
  return insertStyledTextParagraphAfter(body, anchorIndex, textContent);
}

// Inserts a section that can contain both text and code blocks
function insertMixedSectionAfter(body, anchorIndex, label, content) {
  if (isBlank(content)) return anchorIndex;

  body.insertParagraph(anchorIndex + 1, label)
    .setHeading(DocumentApp.ParagraphHeading.HEADING4);

  anchorIndex++;

  const blocks = parseMixedContent(content);

  if (!blocks.length) {
    body.insertParagraph(anchorIndex + 1, content)
      .setHeading(DocumentApp.ParagraphHeading.NORMAL);
    anchorIndex++;
    return anchorIndex;
  }

  blocks.forEach(block => {
    if (block.type === 'text') {
      anchorIndex = insertTextParagraphAfter(body, anchorIndex, block.content);
    } else if (block.type === 'code') {
      anchorIndex = insertCodeBlockAuto(body, anchorIndex, block.content);
    }
  });

  return anchorIndex;
}

// Inserts a simple section with just plain text
function insertSimpleSectionAfter(body, anchorIndex, label, content) {
  if (isBlank(content)) return anchorIndex;

  body.insertParagraph(anchorIndex + 1, label)
    .setHeading(DocumentApp.ParagraphHeading.HEADING4);

  return insertStyledTextParagraphAfter(body, anchorIndex + 1, content || '');
}

/* =========================================================
   SUMMARY
   ========================================================= */

// Counts how many findings exist per severity
function countFindingsBySeverity(findings) {
  const counts = {
    High: 0,
    Medium: 0,
    Low: 0,
    Gas: 0,
    QA: 0
  };

  findings.forEach(f => {
    if (counts[f.severity] !== undefined) {
      counts[f.severity]++;
    }
  });

  return counts;
}

// Builds a singular or plural label for a count
function formatCount(count, singular, plural) {
  if (count === 1) return `${count} ${singular}`;
  return `${count} ${plural}`;
}

// Builds the summary lines shown in the report
function buildSummaryLines(counts) {
  const lines = [];

  if (counts.High > 0) {
    lines.push({
      text: formatCount(counts.High, 'High severity vulnerability', 'High severity vulnerabilities'),
      color: '#ff0000'
    });
  }

  if (counts.Medium > 0) {
    lines.push({
      text: formatCount(counts.Medium, 'Medium severity vulnerability', 'Medium severity vulnerabilities'),
      color: '#ff9900'
    });
  }

  if (counts.Low > 0) {
    lines.push({
      text: formatCount(counts.Low, 'Low severity vulnerability', 'Low severity vulnerabilities'),
      color: '#000000'
    });
  }

  if (counts.Gas > 0) {
    lines.push({
      text: formatCount(counts.Gas, 'Gas Optimisation', 'Gas Optimisations'),
      color: '#000000'
    });
  }

  if (counts.QA > 0) {
    lines.push({
      text: `${counts.QA} QA`,
      color: '#000000'
    });
  }

  return lines;
}

// Updates the severity summary block in the document
function updateSeveritySummary(body, counts) {
  const summaryIndex = findParagraphIndexContaining(body, 'Hashlock found:');

  if (summaryIndex === -1) {
    return 'Summary block "Hashlock found:" not found.';
  }

  const lines = buildSummaryLines(counts);

  let removeIndex = summaryIndex + 1;
  while (removeIndex < body.getNumChildren()) {
    const el = body.getChild(removeIndex);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) break;

    const txt = el.asParagraph().getText().trim();

    if (
      /severity vulnerability/i.test(txt) ||
      /severity vulnerabilities/i.test(txt) ||
      /Gas Optimisation/i.test(txt) ||
      /Gas Optimisations/i.test(txt) ||
      /^\d+\s+QA$/i.test(txt)
    ) {
      body.removeChild(el);
    } else {
      break;
    }
  }

  let insertIndex = summaryIndex + 1;
  lines.forEach(line => {
    const p = body.insertParagraph(insertIndex++, line.text)
      .setHeading(DocumentApp.ParagraphHeading.NORMAL);

    p.editAsText().setForegroundColor(line.color);
  });

  return 'Summary updated successfully.';
}

/* =========================================================
   STYLE HELPERS
   ========================================================= */

// Returns the heading color for a severity section
function getSeverityHeadingColor(severity) {
  const colors = {
    High: '#ff0000',
    Medium: '#ff9900',
    Low: '#000000',
    Gas: '#000000',
    QA: '#000000'
  };

  return colors[severity] || '#000000';
}

// Renumbers findings inside each severity group
function normalizeFindingIds(groups) {
  const prefixMap = {
    High: 'H',
    Medium: 'M',
    Low: 'L',
    Gas: 'G',
    QA: 'Q'
  };

  Object.keys(groups).forEach(severity => {
    let counter = 1;
    groups[severity].forEach(finding => {
      const prefix = prefixMap[severity];
      finding.id = `${prefix}-${String(counter).padStart(2, '0')}`;
      counter++;
    });
  });
}

// Makes sure a heading starts on a new page
function ensureHeadingStartsOnNewPage(body, headingText) {
  const headingIndex = findParagraphIndexContaining(body, headingText);

  if (headingIndex === -1) {
    return `${headingText} heading not found.`;
  }

  if (headingIndex === 0) {
    return `${headingText} is already at the beginning of the document.`;
  }

  const previousElement = body.getChild(headingIndex - 1);

  if (previousElement.getType() === DocumentApp.ElementType.PAGE_BREAK) {
    return `${headingText} already starts on a new page.`;
  }

  body.insertPageBreak(headingIndex);

  return `${headingText} moved to a new page.`;
}

// Applies the page-break rule to multiple main headings
function ensureMainHeadingsStartOnNewPage(body, headingTexts) {
  const results = [];

  for (let i = 0; i < headingTexts.length; i++) {
    const result = ensureHeadingStartsOnNewPage(body, headingTexts[i]);
    results.push(result);
  }

  return results.join(' ');
}

/* =========================================================
   MAIN WRITER
   ========================================================= */

// Writes structured findings into the audit report
function writeFindingsToReport(inputText) {
  try {
    const findings = parseFindings(inputText);
    if (!findings.length) return 'No findings found.';

    const errors = validateFindings(findings);
    if (errors.length) return 'Errors:\n\n' + errors.join('\n');

    const body = DocumentApp.getActiveDocument().getBody();
    const counts = countFindingsBySeverity(findings);

    let startIndex = findMarkerIndex(body, '[[AUDIT_FINDINGS_START]]');
    let endIndex = findMarkerIndex(body, '[[AUDIT_FINDINGS_END]]');

    if (startIndex === -1 || endIndex === -1) {
      return 'Markers not found.';
    }

    if (endIndex <= startIndex) {
      return 'Markers are in an invalid order.';
    }

    // Clears the old findings block between the markers
    for (let i = endIndex - 1; i > startIndex; i--) {
      body.removeChild(body.getChild(i));
    }

    startIndex = findMarkerIndex(body, '[[AUDIT_FINDINGS_START]]');

    const groups = {
      High: [],
      Medium: [],
      Low: [],
      Gas: [],
      QA: []
    };

    findings.forEach(f => groups[f.severity].push(f));

    normalizeFindingIds(groups);

    let cursor = startIndex;
    let firstFindingOverall = true;

    Object.keys(groups).forEach(severity => {
      if (!groups[severity].length) return;

      groups[severity].forEach((finding, indexInSeverity) => {
        // Adds a page break before every finding except the first one
        if (!firstFindingOverall) {
          body.insertPageBreak(cursor + 1);
          cursor++;
        }

        firstFindingOverall = false;

        // Adds the severity heading once per severity group
        if (indexInSeverity === 0) {
          const severityParagraph = body.insertParagraph(cursor + 1, severity)
            .setHeading(DocumentApp.ParagraphHeading.HEADING2);

          const severityText = severityParagraph.editAsText();
          severityText.setForegroundColor(getSeverityHeadingColor(severity));
          cursor++;
        }

        const parsedTitle = parseInlineCodeRanges(`[${finding.id}] ${finding.title}`);
        const fullTitle = parsedTitle.text;

        const titleParagraph = body.insertParagraph(cursor + 1, fullTitle)
          .setHeading(DocumentApp.ParagraphHeading.HEADING3);

        const titleText = titleParagraph.editAsText();

        const titleParts = finding.title.split('`')[0].split(' - ');
        const rawTitle = finding.title.replace(/`[^`]*`/g, match => match.slice(1, -1));
        const rawTitleParts = rawTitle.split(' - ');
        const functionPart = rawTitleParts[0];
        const restPart = rawTitleParts.slice(1).join(' - ');

        const idPrefix = `[${finding.id}] `;
        const startFunction = fullTitle.indexOf(functionPart);
        const endFunction = startFunction !== -1 ? startFunction + functionPart.length - 1 : -1;

        const restSearch = restPart ? ' - ' + restPart : '';
        const startRest = restSearch ? fullTitle.indexOf(restSearch) : -1;
        const endRest = startRest !== -1 ? startRest + restSearch.length - 1 : -1;

        // Styles the function/method part of the title
        if (startFunction !== -1) {
          titleText.setForegroundColor(startFunction, endFunction, '#7951a1');
          titleText.setBold(startFunction, endFunction, true);
        }

        // Styles the remaining title text
        if (startRest !== -1) {
          titleText.setForegroundColor(startRest, endRest, '#00c0a3');
          titleText.setBold(startRest, endRest, false);
        }

        // Applies inline code highlight to backtick ranges in the title
        parsedTitle.ranges.forEach(function (range) {
          if (range.start >= 0 && range.end >= range.start) {
            titleText.setFontFamily(range.start, range.end, 'Maven Pro');
            titleText.setFontSize(range.start, range.end, 13);
            titleText.setBackgroundColor(range.start, range.end, '#eeeeee');
            titleText.setForegroundColor(range.start, range.end, '#000000');
            titleText.setBold(range.start, range.end, false);
          }
        });

        cursor++;

        cursor = insertSimpleSectionAfter(body, cursor, 'Description', finding.description);
        cursor = insertMixedSectionAfter(body, cursor, 'Vulnerability Details', finding.vulnerability_details);
        cursor = insertMixedSectionAfter(body, cursor, 'Proof of Concept', finding.proof_of_concept);
        cursor = insertSimpleSectionAfter(body, cursor, 'Impact', finding.impact);
        cursor = insertMixedSectionAfter(body, cursor, 'Recommendation', finding.recommendation);
        cursor = insertSimpleSectionAfter(body, cursor, 'Status', finding.status);

        body.insertParagraph(cursor + 1, '');
        cursor++;
      });
    });

    const summaryResult = updateSeveritySummary(body, counts);

    return `Success: ${findings.length} findings inserted. ${summaryResult}`;
  } catch (e) {
    return 'Error: ' + e.message;
  }
}

/* =========================================================
   MARKDOWN WRITER
   ========================================================= */

// Converts markdown findings and writes them into the report
function writeMarkdownToReport(markdown) {
  try {
    const findings = parseMarkdownFindings(markdown);

    if (!findings.length) {
      return 'No findings found in markdown.';
    }

    const errors = validateFindings(findings);
    if (errors.length) {
      return 'Markdown validation errors:\n\n' + errors.join('\n');
    }

    const structuredText = findingsToStructuredText(findings);
    return writeFindingsToReport(structuredText);
  } catch (e) {
    return 'Error writing markdown to report: ' + e.message;
  }
}

/* =========================================================
   CONTRACT WRITER
   ========================================================= */

// Writes contract names into the target scope table
function writeContracts(contractsText) {
  try {
    var contracts = String(contractsText || '')
      .split('\n')
      .map(function (c) { return c.trim(); })
      .filter(function (c) { return c.length > 0; });

    if (!contracts.length) return 'No contracts found.';

    var body = DocumentApp.getActiveDocument().getBody();
    var tables = body.getTables();

    // Finds the table that contains the "Contract 1" row
    var targetTable = null;
    var contractRowIndex = -1;

    for (var t = 0; t < tables.length; t++) {
      var table = tables[t];
      for (var r = 0; r < table.getNumRows(); r++) {
        var row = table.getRow(r);
        if (row.getNumCells() < 2) continue;

        if (row.getCell(0).getText().trim() === 'Contract 1') {
          targetTable = table;
          contractRowIndex = r;
          break;
        }
      }
      if (targetTable) break;
    }

    if (!targetTable) {
      return 'Could not find "Contract 1" row.';
    }

    // Counts how many contract rows already exist
    var existingCount = 0;
    for (var r = contractRowIndex; r < targetTable.getNumRows(); r++) {
      var row = targetTable.getRow(r);
      if (row.getNumCells() < 2) break;

      var label = row.getCell(0).getText().trim();
      if (/^Contract \d+$/.test(label)) {
        existingCount++;
      } else {
        break;
      }
    }

    // Fills existing rows first
    for (var i = 0; i < existingCount && i < contracts.length; i++) {
      var row = targetTable.getRow(contractRowIndex + i);
      row.getCell(0).setText('Contract ' + (i + 1));
      row.getCell(1).setText(contracts[i]);
    }

    // Adds new rows if more contracts are needed
    for (var i = existingCount; i < contracts.length; i++) {
      var newRow = targetTable.insertTableRow(contractRowIndex + i);
      newRow.appendTableCell('Contract ' + (i + 1));
      newRow.appendTableCell(contracts[i]);
    }

    // Removes extra rows if fewer contracts were provided
    for (var i = contracts.length; i < existingCount; i++) {
      targetTable.removeRow(contractRowIndex + contracts.length);
    }

    // Reapplies the standard style to all used rows
    for (var i = 0; i < contracts.length; i++) {
      var row = targetTable.getRow(contractRowIndex + i);

      var leftCell = row.getCell(0);
      var rightCell = row.getCell(1);

      leftCell.setText('Contract ' + (i + 1));
      rightCell.setText(contracts[i]);

      applyContractCellStyle(leftCell, true, 'Maven Pro', 12);
      applyContractCellStyle(rightCell, false, 'Maven Pro', 12);
    }

    // Applies a standard border style to the table
    targetTable.setBorderWidth(1);
    targetTable.setBorderColor('#444444');

    return 'Done: ' + contracts.length + ' contract(s) written.';
  } catch (e) {
    return 'Error: ' + e.message;
  }
}

/* =========================================================
   CONTRACT STYLE HELPER
   ========================================================= */

// Applies a clean text style to a contract table cell
function applyContractCellStyle(cell, isBold, fontFamily, fontSize) {
  var text = cell.editAsText();

  text.setBold(false);
  text.setFontFamily(fontFamily);
  text.setFontSize(fontSize);

  if (isBold) {
    text.setBold(true);
  }
}

/* =========================================================
   INLINE CODE STYLING
   ========================================================= */

// Finds inline code marked with backticks and tracks its ranges
function parseInlineCodeRanges(text) {
  var source = String(text || '');
  var cleanText = '';
  var ranges = [];

  var i = 0;
  while (i < source.length) {
    if (source[i] === '`') {
      var end = source.indexOf('`', i + 1);

      // If there is no closing backtick, treat it as regular text
      if (end === -1) {
        cleanText += source[i];
        i++;
        continue;
      }

      var codeText = source.slice(i + 1, end);
      var startIndex = cleanText.length;

      cleanText += codeText;

      if (codeText.length > 0) {
        ranges.push({
          start: startIndex,
          end: cleanText.length - 1
        });
      }

      i = end + 1;
      continue;
    }

    cleanText += source[i];
    i++;
  }

  return {
    text: cleanText,
    ranges: ranges
  };
}

// Applies inline code styling to marked text ranges
function applyInlineCodeStyle(textElement, ranges) {
  if (!ranges || !ranges.length) return;

  ranges.forEach(function (range) {
    if (range.start >= 0 && range.end >= range.start) {
      textElement.setFontFamily(range.start, range.end, 'Roboto Mono');
      textElement.setFontSize(range.start, range.end, 10);
      textElement.setBackgroundColor(range.start, range.end, '#eeeeee');
      textElement.setForegroundColor(range.start, range.end, '#000000');
      textElement.setBold(range.start, range.end, false);
    }
  });
}

// Inserts a paragraph and styles any inline code inside it
function insertStyledTextParagraphAfter(body, anchorIndex, textContent) {
  var parsed = parseInlineCodeRanges(textContent || '');

  var paragraph = body.insertParagraph(anchorIndex + 1, parsed.text)
    .setHeading(DocumentApp.ParagraphHeading.NORMAL);

  var text = paragraph.editAsText();
  applyInlineCodeStyle(text, parsed.ranges);

  return anchorIndex + 1;
}
