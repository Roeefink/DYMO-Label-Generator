import React, { useState } from 'react';
import * as XLSX from 'xlsx';

interface ParsedItem {
  productId: string;
  description: string;
  amount: number;
}

export default function App() {
  const [fileName, setFileName] = useState<string>('');
  const [parsedData, setParsedData] = useState<ParsedItem[]>([]);
  const [totalLabels, setTotalLabels] = useState<number>(0);
  const [error, setError] = useState<string | null>(null);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    setError(null);
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const buffer = event.target?.result;
        const workbook = XLSX.read(buffer, { type: 'binary' });

        // Grab first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Convert sheet to JSON rows
        const rawRows: Record<string, any>[] = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

        if (!rawRows.length) {
          setError('The uploaded file appears to be empty.');
          return;
        }

        // Find relevant column keys dynamically
        const keys = Object.keys(rawRows[0]);
        const idKey = keys.find(k => /id|sku|part|item.*no/i.test(k)) || keys[0];
        const descKey = keys.find(k => /desc/i.test(k)) || keys[1];
        const amountKey = keys.find(k => /amount|qty|quantity|count/i.test(k)) || keys[2];

        let labelCount = 0;
        const cleanedItems: ParsedItem[] = [];

        rawRows.forEach((row) => {
          const productId = String(row[idKey] ?? '').trim();
          const description = String(row[descKey] ?? '').trim();
          const amount = parseInt(row[amountKey], 10) || 0;

          if (productId && amount > 0) {
            cleanedItems.push({ productId, description, amount });
            labelCount += amount;
          }
        });

        if (cleanedItems.length === 0) {
          setError('Could not identify valid items with amounts greater than 0.');
          return;
        }

        setParsedData(cleanedItems);
        setTotalLabels(labelCount);
      } catch (err) {
        console.error(err);
        setError('Failed to parse Excel file. Please verify the format.');
      }
    };

    reader.readAsBinaryString(file);
  };

  const handleExport = () => {
    if (!parsedData.length) return;

    // Multiply rows by the amount specified
    const flattenedRows: { 'Product ID': string; 'Description': string }[] = [];

    parsedData.forEach((item) => {
      for (let i = 0; i < item.amount; i++) {
        flattenedRows.push({
          'Product ID': item.productId,
          'Description': item.description,
        });
      }
    });

    // Create workbook and export
    const worksheet = XLSX.utils.json_to_sheet(flattenedRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'DYMO_Labels');

    const outputName = fileName.replace(/\.[^/.]+$/, '') + '_dymo_ready.xlsx';
    XLSX.writeFile(workbook, outputName);
  };

  return (
    <main style={{ maxWidth: 640, margin: '40px auto', fontFamily: 'system-ui, sans-serif', padding: 24 }}>
      <h2>DYMO Label Generator</h2>
      <p style={{ color: '#555' }}>
        Upload your inventory Excel sheet. The app will extract the Product ID and Description, then duplicate each item by its Amount so DYMO prints the exact quantity.
      </p>

      <div style={{ border: '2px dashed #ccc', padding: 28, textAlign: 'center', borderRadius: 8, margin: '20px 0' }}>
        <input
          type="file"
          accept=".xlsx, .xls, .csv"
          onChange={handleFileUpload}
          id="file-input"
          style={{ display: 'none' }}
        />
        <label
          htmlFor="file-input"
          style={{ cursor: 'pointer', padding: '10px 18px', background: '#0070c9', color: '#fff', borderRadius: 6, fontWeight: 500 }}
        >
          Select Inventory File
        </label>
        {fileName && <p style={{ marginTop: 14, fontSize: 14, color: '#333' }}>Selected: <strong>{fileName}</strong></p>}
      </div>

      {error && <div style={{ color: '#c00', marginBottom: 16 }}>{error}</div>}

      {parsedData.length > 0 && (
        <div style={{ background: '#f8f9fa', border: '1px solid #e2e8f0', borderRadius: 8, padding: 16 }}>
          <p style={{ margin: 0, marginBottom: 12 }}>
            Found <strong>{parsedData.length}</strong> unique products | Total labels to print: <strong>{totalLabels}</strong>
          </p>

          <button
            onClick={handleExport}
            style={{
              width: '100%',
              padding: '12px',
              backgroundColor: '#34c759',
              color: '#fff',
              border: 'none',
              borderRadius: 6,
              fontSize: 16,
              fontWeight: 600,
              cursor: 'pointer'
            }}
          >
            Download DYMO Ready Sheet (.xlsx)
          </button>
        </div>
      )}
    </main>
  );
}
