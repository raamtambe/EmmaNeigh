const fs = require('fs');
const path = require('path');
const { PDFDocument } = require('pdf-lib');

/**
 * Pure JavaScript PDF processor - no Python required
 * Uses pdf-lib for PDF manipulation (no DOMMatrix dependency)
 */

// Normalize signer name
function normalizeName(name) {
  return name
    .toUpperCase()
    .replace(/[.,]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

// Check if name is probably a person (not an entity)
function isProbablePerson(name) {
  const entityTerms = ['LLC', 'INC', 'CORP', 'CORPORATION', 'LP', 'LLP', 'TRUST', 'HOLDINGS', 'PARTNERS', 'FUND', 'CAPITAL'];
  const upperName = name.toUpperCase();
  if (entityTerms.some(term => upperName.includes(term))) {
    return false;
  }
  const words = name.split(' ').filter(w => w.length > 0);
  return words.length >= 2 && words.length <= 4;
}

// Simple text extraction from PDF buffer using basic parsing
// This is a simplified approach that looks for text patterns
function extractTextFromPdfBuffer(buffer) {
  try {
    const content = buffer.toString('binary');

    // Look for text between BT and ET markers (basic PDF text extraction)
    const textMatches = [];
    const streamRegex = /stream[\r\n]+([\s\S]*?)[\r\n]+endstream/g;
    let match;

    while ((match = streamRegex.exec(content)) !== null) {
      const streamContent = match[1];
      // Look for text showing operators
      const textRegex = /\(([^)]+)\)/g;
      let textMatch;
      while ((textMatch = textRegex.exec(streamContent)) !== null) {
        textMatches.push(textMatch[1]);
      }
    }

    return textMatches.join(' ');
  } catch (err) {
    return '';
  }
}

// Extract signers from page text
function extractSignersFromText(text) {
  const lines = text.split(/[\n\r]+/).map(l => l.trim()).filter(l => l);
  const signers = new Set();

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    if (line.toUpperCase().includes('BY:') || line.toUpperCase().includes('BY :')) {
      // Tier 1: Look for explicit Name: field
      let foundName = false;
      for (let j = 1; j <= 6 && i + j < lines.length; j++) {
        const candidate = lines[i + j];
        if (candidate.toUpperCase().startsWith('NAME:') || candidate.toUpperCase().startsWith('NAME :')) {
          const name = candidate.split(':')[1]?.trim();
          if (name && name.length > 2) {
            signers.add(normalizeName(name));
            foundName = true;
          }
          break;
        }
      }

      // Tier 2: Look for probable person name
      if (!foundName) {
        for (let j = 1; j <= 6 && i + j < lines.length; j++) {
          const candidate = normalizeName(lines[i + j]);
          if (candidate && candidate.length > 3 && isProbablePerson(candidate)) {
            signers.add(candidate);
            break;
          }
        }
      }
    }
  }

  return Array.from(signers);
}

// Process signature packets
async function processSignaturePackets(inputDir, progressCallback) {
  // Find all PDF files
  const files = fs.readdirSync(inputDir).filter(f => f.toLowerCase().endsWith('.pdf'));

  if (files.length === 0) {
    throw new Error('No PDF files found in the selected folder.');
  }

  progressCallback('scanning', 0, `Found ${files.length} PDF files`);

  // Collect signature data
  const signerDocPages = {}; // { signerName: { docName: [pages] } }

  for (let i = 0; i < files.length; i++) {
    const filename = files[i];
    const percent = Math.round((i / files.length) * 50);
    progressCallback('scanning', percent, `Scanning ${filename}`);

    const pdfPath = path.join(inputDir, filename);
    const dataBuffer = fs.readFileSync(pdfPath);

    try {
      const pdfDoc = await PDFDocument.load(dataBuffer, { ignoreEncryption: true });
      const pageCount = pdfDoc.getPageCount();

      // Extract text and find signers
      const text = extractTextFromPdfBuffer(dataBuffer);
      const signers = extractSignersFromText(text);

      // If we found signers, associate all pages with them
      // (In a real scenario, you'd want page-level detection)
      if (signers.length > 0) {
        for (const signer of signers) {
          if (!signerDocPages[signer]) {
            signerDocPages[signer] = {};
          }
          if (!signerDocPages[signer][filename]) {
            signerDocPages[signer][filename] = [];
          }
          // For now, include all pages from documents where signer was found
          for (let p = 1; p <= pageCount; p++) {
            if (!signerDocPages[signer][filename].includes(p)) {
              signerDocPages[signer][filename].push(p);
            }
          }
        }
      }
    } catch (err) {
      progressCallback('warning', percent, `Error processing ${filename}: ${err.message}`);
    }
  }

  const signerNames = Object.keys(signerDocPages);

  if (signerNames.length === 0) {
    throw new Error('No signature pages detected. Make sure PDFs contain "BY:" and "Name:" fields.');
  }

  progressCallback('extracting', 50, `Creating packets for ${signerNames.length} signers`);

  // Create output directory
  const outputBase = path.join(inputDir, 'signature_packets_output');
  const outputPdfDir = path.join(outputBase, 'packets');
  fs.mkdirSync(outputPdfDir, { recursive: true });

  const packetsCreated = [];

  // Create signature packets
  for (let i = 0; i < signerNames.length; i++) {
    const signerName = signerNames[i];
    const docPages = signerDocPages[signerName];
    const percent = 50 + Math.round((i / signerNames.length) * 50);
    progressCallback('extracting', percent, `Producing signature packet for ${signerName}`);

    try {
      const packetPdf = await PDFDocument.create();

      for (const [docName, pages] of Object.entries(docPages)) {
        const srcPath = path.join(inputDir, docName);
        const srcBuffer = fs.readFileSync(srcPath);
        const srcPdf = await PDFDocument.load(srcBuffer, { ignoreEncryption: true });

        // Copy pages (convert to 0-indexed)
        const pageIndices = pages.map(p => p - 1);
        const copiedPages = await packetPdf.copyPages(srcPdf, pageIndices);
        for (const page of copiedPages) {
          packetPdf.addPage(page);
        }
      }

      if (packetPdf.getPageCount() > 0) {
        const pdfBytes = await packetPdf.save();
        const outputPath = path.join(outputPdfDir, `signature_packet - ${signerName}.pdf`);
        fs.writeFileSync(outputPath, pdfBytes);

        packetsCreated.push({
          name: signerName,
          pages: packetPdf.getPageCount(),
          path: outputPath
        });
      }
    } catch (err) {
      progressCallback('warning', percent, `Error creating packet for ${signerName}: ${err.message}`);
    }
  }

  progressCallback('complete', 100, `Created ${packetsCreated.length} signature packets`);

  return {
    success: true,
    packetsCreated: packetsCreated.length,
    packets: packetsCreated,
    outputPath: outputBase
  };
}

// Create execution version
async function createExecutionVersion(originalPath, signedPath, insertAfter, progressCallback) {
  progressCallback('loading', 10, 'Loading original document...');

  const originalBuffer = fs.readFileSync(originalPath);
  const originalPdf = await PDFDocument.load(originalBuffer, { ignoreEncryption: true });
  const originalPageCount = originalPdf.getPageCount();

  progressCallback('loading', 30, 'Loading signed document...');

  const signedBuffer = fs.readFileSync(signedPath);
  const signedPdf = await PDFDocument.load(signedBuffer, { ignoreEncryption: true });
  const signedPageCount = signedPdf.getPageCount();

  progressCallback('merging', 50, `Merging ${signedPageCount} signed pages...`);

  const resultPdf = await PDFDocument.create();

  // Determine insertion point
  let insertPoint = insertAfter;
  if (insertPoint < 0 || insertPoint >= originalPageCount) {
    insertPoint = originalPageCount;
  }

  // Copy pages before insertion point
  if (insertPoint > 0) {
    progressCallback('merging', 60, 'Adding pages before signature pages...');
    const beforePages = await resultPdf.copyPages(originalPdf,
      Array.from({ length: insertPoint }, (_, i) => i)
    );
    for (const page of beforePages) {
      resultPdf.addPage(page);
    }
  }

  // Insert signed pages
  progressCallback('merging', 75, 'Inserting signed pages...');
  const signedPages = await resultPdf.copyPages(signedPdf,
    Array.from({ length: signedPageCount }, (_, i) => i)
  );
  for (const page of signedPages) {
    resultPdf.addPage(page);
  }

  // Copy remaining pages
  if (insertPoint < originalPageCount) {
    progressCallback('merging', 90, 'Adding remaining pages...');
    const afterPages = await resultPdf.copyPages(originalPdf,
      Array.from({ length: originalPageCount - insertPoint }, (_, i) => i + insertPoint)
    );
    for (const page of afterPages) {
      resultPdf.addPage(page);
    }
  }

  // Generate output filename
  const originalBasename = path.basename(originalPath);
  const nameWithoutExt = originalBasename.replace(/\.pdf$/i, '')
    .replace(/_Clean$/, '')
    .replace(/_Without_Sigs$/, '')
    .replace(/_Unsigned$/, '')
    .replace(/_Draft$/, '');

  const outputFilename = `${nameWithoutExt} (Execution Version).pdf`;
  const tempDir = require('os').tmpdir();
  const outputPath = path.join(tempDir, outputFilename);

  progressCallback('saving', 95, 'Saving execution version...');
  const pdfBytes = await resultPdf.save();
  fs.writeFileSync(outputPath, pdfBytes);

  progressCallback('complete', 100, 'Execution version created successfully!');

  return {
    success: true,
    outputPath,
    outputFilename,
    originalPages: originalPageCount,
    signedPages: signedPageCount,
    totalPages: resultPdf.getPageCount()
  };
}

module.exports = {
  processSignaturePackets,
  createExecutionVersion
};
