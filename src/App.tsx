import React, { useState } from 'react';
import * as XLSX from 'xlsx';

interface ParsedItem {
  topLine: string;
  bottomLine: string;
  amount: number;
  category: DeviceCategory;
  sourceKey: string;
  needsProcessor: boolean;
}

type DeviceCategory = 'iPad' | 'Apple Watch' | 'Mac';

const deviceCategories: DeviceCategory[] = ['iPad', 'Apple Watch', 'Mac'];

const normalizeHeader = (header: string) => header.toLowerCase().replace(/[^a-z0-9]/g, '');

const parseAmount = (value: unknown) => {
  if (typeof value === 'number') return Math.floor(value);
  const normalized = String(value ?? '').replace(/,/g, '').trim();
  const amount = Number.parseFloat(normalized);
  return Number.isFinite(amount) ? Math.floor(amount) : 0;
};

const getDeviceCategory = (value: string): DeviceCategory | null => {
  if (/charger|cable|adapter|pencil|airpods|case|mouse|keyboard|strap|battery|earpods|\baccessor(?:y|ies)\b/i.test(value)) return null;
  if (/\bipad\b/i.test(value)) return 'iPad';
  if (/apple\s+watch/i.test(value)) return 'Apple Watch';
  if (/macbook|\bmbp\b|\bmba\b|\bmbn\b/i.test(value)) return 'Mac';
  return null;
};

const formatMacLabel = (value: string) => {
  const familyMatch = value.match(/\b(MBP|MBA|MBN)\s+(\d+)/i);
  const family = familyMatch?.[1].toUpperCase() || 'Mac';
  const size = familyMatch?.[2] || '';
  const processor = value.match(/\b(?:A\d+\s*Pro|M\d+(?:\s+(?:Pro|Max|Ultra))?)\b/i)?.[0] || '';
  const coreMatch = value.match(/\b(\d+)C\s*\/\s*(\d+)C\s*GPU\b/i);
  const normalizeCapacity = (value: string, unit: string) => {
    const normalizedUnit = unit.toUpperCase();
    return `${value}${normalizedUnit === 'G' ? 'GB' : normalizedUnit === 'T' ? 'TB' : normalizedUnit}`;
  };
  const slashCapacity = value.match(/\b(\d+(?:\.\d+)?)\s*(GB|G)?\s*\/\s*(\d+(?:\.\d+)?)\s*(GB|TB|G|T)\b/i);
  const ramAndStorage = slashCapacity
    ? [
      normalizeCapacity(slashCapacity[1], slashCapacity[2] || 'GB'),
      normalizeCapacity(slashCapacity[3], slashCapacity[4]),
    ]
    : [...value.matchAll(/\b(\d+(?:\.\d+)?)\s*(GB|TB|G|T)\b/gi)].map(match => {
      return normalizeCapacity(match[1], match[2]);
    });
  const ram = ramAndStorage[0] || '';
  const storage = ramAndStorage[1] || '';
  const colorSource = family === 'MBN'
    ? value.match(/\b(IND|BLS|CIT|SLV|indigo|blue|blush|pink|citrus|yellow|silver)\b/i)?.[1]
    : value.match(/\b(SB|SLV|STL|SKY|MDN|IND|BLS|CIT)\b/i)?.[1];
  const color = colorSource
    ? ({ indigo: 'IND', blue: 'IND', blush: 'BLS', pink: 'BLS', citrus: 'CIT', yellow: 'CIT', silver: 'SLV' }[colorSource.toLowerCase()] || colorSource.toUpperCase())
    : '';
  const topLine = [family, size, processor].filter(Boolean).join(' ');
  const memory = ram && storage ? `${ram}/${storage}` : ram || storage;
  const cores = coreMatch ? `${coreMatch[1]}C CPU/${coreMatch[2]}C GPU` : '';
  const bottomLine = [color, memory, cores].filter(Boolean).join(' - ');
  return { topLine, bottomLine };
};

const formatProductLabel = (productName: string, details: string, processorOverride = '') => {
  const source = `${productName} ${details}`.replace(/\s+/g, ' ').trim();
  const normalized = source.replace(/[‐‑‒–—]/g, '-');
  const lowerSource = normalized.toLowerCase();

  if (/macbook\s+pro|\bmbp\b/i.test(normalized)) {
    return { ...formatMacLabel(normalized.replace(/macbook\s+pro/ig, 'MBP')), needsProcessor: false };
  }
  if (/macbook\s+air|\bmba\b/i.test(normalized)) {
    return { ...formatMacLabel(normalized.replace(/macbook\s+air/ig, 'MBA')), needsProcessor: false };
  }
  if (/macbook\s+neo|\bneo\b|\bmbn\b/i.test(normalized)) {
    return { ...formatMacLabel(normalized.replace(/macbook\s+neo/ig, 'MBN')), needsProcessor: false };
  }
  if (/apple\s+watch/i.test(normalized)) {
    const watchDetails = /^apple\s+watch$/i.test(productName.trim()) ? details : normalized.replace(/apple\s+watch/ig, '');
    return { topLine: 'Apple Watch', bottomLine: watchDetails.replace(/^[A-Z0-9]+(?:\/A)?\s+/i, '').trim(), needsProcessor: false };
  }
  if (/\bipad\b/i.test(normalized)) {
    const family = /ipad\s+pro/i.test(normalized) ? 'iPad Pro' : /ipad\s+air/i.test(normalized) ? 'iPad Air' : /ipad\s+mini/i.test(normalized) ? 'iPad mini' : 'iPad';
    const detectedSize = normalized.match(/(\d+(?:\.\d+)?)\s*-?\s*inch/i)?.[1];
    const size = detectedSize || (family === 'iPad Air' ? '11' : family === 'iPad' ? '11' : '');
    const processor = processorOverride.trim() || normalized.match(/\b(?:M\d|A\d+(?:\s*Pro)?)\b/i)?.[0]
      || (family === 'iPad mini' ? 'A17 Pro' : '')
      || (size === '11' && !/air|pro|mini/i.test(normalized) ? 'A16' : '');
    const topLine = family === 'iPad mini'
      ? [family, processor].filter(Boolean).join(' ')
      : [family, size, processor].filter(Boolean).join(' - ');
    const color = normalized.match(/\b(black|blue|purple|pink|silver|starlight|space\s+gray|space\s+grey|gold|yellow)\b/i)?.[0];
    const capacity = normalized.match(/\b\d+(?:\.\d+)?\s*(?:GB|TB|G|T)\b/i)?.[0].replace(/\s+/g, '').replace(/G$/, 'GB').replace(/T$/, 'TB');
    const connectivity = /wi-?fi\s*\+\s*cell|cellular/i.test(lowerSource) ? 'WiFi + Cell' : /wi-?fi/i.test(lowerSource) ? 'WiFi' : '';
    const bottomLine = [color, capacity && `${color ? '- ' : ''}${capacity}`, connectivity].filter(Boolean).join(' ').trim();
    return { topLine, bottomLine, needsProcessor: !processor };
  }

  return { topLine: productName.trim(), bottomLine: details.trim(), needsProcessor: false };
};

export default function App() {
  const [fileName, setFileName] = useState<string>('');
  const [parsedData, setParsedData] = useState<ParsedItem[]>([]);
  const [pendingData, setPendingData] = useState<ParsedItem[]>([]);
  const [selectedCategories, setSelectedCategories] = useState<DeviceCategory[]>(deviceCategories);
  const [totalLabels, setTotalLabels] = useState<number>(0);
  const [outputFileName, setOutputFileName] = useState<string>('');
  const [processorInputs, setProcessorInputs] = useState<Record<string, string>>({});
  const [error, setError] = useState<string | null>(null);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    setError(null);
    const file = e.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    setOutputFileName(`${file.name.replace(/\.[^/.]+$/, '')}_dymo_ready`);
    setProcessorInputs({});

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const buffer = event.target?.result;
        const workbook = XLSX.read(buffer, { type: 'array' });

        // Grab first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Read rows directly so reports with a title row before the headers still work.
        const rawRows = XLSX.utils.sheet_to_json<unknown[]>(worksheet, { header: 1, defval: '' });

        if (!rawRows.length) {
          setError('The uploaded file appears to be empty.');
          return;
        }

        const idPatterns = [/productid/, /sku/, /partnumber/, /itemnumber/, /^item$/, /^id$/, /topline/];
        const descPatterns = [/description/, /productname/, /itemname/, /name/, /bottomline/];
        const amountPatterns = [/amount/, /qty/, /quantity/, /count/, /onhand/, /instock/, /available/, /stock/, /inventory/];
        const headerIndex = rawRows.findIndex((row, index) => {
          if (index > 20) return false;
          const normalizedRow = row.map(value => normalizeHeader(String(value ?? '')));
          const matches = (patterns: RegExp[]) => normalizedRow.some(value => patterns.some(pattern => pattern.test(value)));
          return Number(matches(idPatterns)) + Number(matches(descPatterns)) + Number(matches(amountPatterns)) >= 2;
        });
        const actualHeaderIndex = headerIndex >= 0 ? headerIndex : 0;
        const headers = rawRows[actualHeaderIndex].map((value, index) => String(value || `Column ${index + 1}`));
        const findColumn = (patterns: RegExp[]) => headers.findIndex(header => patterns.some(pattern => pattern.test(normalizeHeader(header))));
        const idColumn = findColumn(idPatterns);
        const descColumn = findColumn(descPatterns);
        const amountColumn = findColumn(amountPatterns);
        const dataRows = rawRows.slice(actualHeaderIndex + 1);
        const getValue = (row: unknown[], column: number, fallbackColumn: number) => row[column >= 0 ? column : fallbackColumn];

        let labelCount = 0;
        const cleanedItems: ParsedItem[] = [];

        dataRows.forEach((row, rowIndex) => {
          const productId = String(getValue(row, idColumn, 0) ?? '').trim();
          const description = String(getValue(row, descColumn, 1) ?? '').trim();
          const category = getDeviceCategory(`${productId} ${description}`);
          const label = formatProductLabel(productId, description);
          const amount = amountColumn >= 0 ? parseAmount(getValue(row, amountColumn, 2)) : (productId ? 1 : 0);

          if (productId && category && amount > 0) {
            cleanedItems.push({ ...label, amount, category, sourceKey: `${actualHeaderIndex + rowIndex + 1}`, });
            labelCount += amount;
          }
        });

        if (cleanedItems.length === 0) {
          setError('No iPad, Apple Watch, or Mac items were found in the uploaded file.');
          return;
        }

        setPendingData(cleanedItems);
        setParsedData([]);
        setTotalLabels(labelCount);
        setSelectedCategories(deviceCategories);
      } catch (err) {
        console.error(err);
        setError('Failed to parse Excel file. Please verify the format.');
      }
    };

    reader.readAsArrayBuffer(file);
  };

  const toggleCategory = (category: DeviceCategory) => {
    setSelectedCategories((current) => current.includes(category)
      ? current.filter((selected) => selected !== category)
      : [...current, category]);
  };

  const applyFilter = () => {
    const filteredItems = pendingData.filter(item => selectedCategories.includes(item.category));
    setParsedData(filteredItems);
    setTotalLabels(filteredItems.reduce((total, item) => total + item.amount, 0));
    setError(filteredItems.length ? null : 'Select at least one device type with matching items.');
  };

  const updateProcessor = (sourceKey: string, processor: string) => {
    setProcessorInputs((current) => ({ ...current, [sourceKey]: processor }));
  };

  const confirmProcessor = (sourceKey: string, processor: string) => {
    if (!processor.trim()) return;
    setParsedData((items) => items.map((item) => item.sourceKey === sourceKey
      ? { ...item, topLine: `${item.topLine} ${processor.trim()}`.trim(), needsProcessor: false }
      : item));
  };

  const handleExport = () => {
    if (!parsedData.length) return;

    // Multiply rows by the amount specified
    const flattenedRows: { TopLine: string; BottomLine: string }[] = [];

    parsedData.forEach((item) => {
      for (let i = 0; i < item.amount; i++) {
        flattenedRows.push({
          TopLine: item.topLine,
          BottomLine: item.bottomLine,
        });
      }
    });

    // Create workbook and export
    const worksheet = XLSX.utils.json_to_sheet(flattenedRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'DYMO_Labels');

    const cleanName = outputFileName.trim() || 'dymo_ready';
    const outputName = /\.xlsx$/i.test(cleanName) ? cleanName : `${cleanName}.xlsx`;
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

      {pendingData.length > 0 && parsedData.length === 0 && (
        <div style={{ background: '#f8f9fa', border: '1px solid #e2e8f0', borderRadius: 8, padding: 16, marginBottom: 16 }}>
          <p style={{ marginTop: 0, fontWeight: 600 }}>Which devices should be included?</p>
          {deviceCategories.map((category) => {
            const count = pendingData
              .filter(item => item.category === category)
              .reduce((total, item) => total + item.amount, 0);
            return (
              <label key={category} style={{ display: 'block', margin: '10px 0', cursor: 'pointer' }}>
                <input
                  type="checkbox"
                  checked={selectedCategories.includes(category)}
                  onChange={() => toggleCategory(category)}
                  style={{ marginRight: 8 }}
                />
                {category} ({count} labels)
              </label>
            );
          })}
          <button
            onClick={applyFilter}
            style={{ width: '100%', padding: '12px', marginTop: 10, backgroundColor: '#0070c9', color: '#fff', border: 'none', borderRadius: 6, fontSize: 16, fontWeight: 600, cursor: 'pointer' }}
          >
            Apply Filter
          </button>
        </div>
      )}

      {parsedData.length > 0 && (
        <div style={{ background: '#f8f9fa', border: '1px solid #e2e8f0', borderRadius: 8, padding: 16 }}>
          <p style={{ margin: 0, marginBottom: 12 }}>
            Found <strong>{parsedData.length}</strong> unique products | Total labels to print: <strong>{totalLabels}</strong>
          </p>

          {parsedData.some(item => item.needsProcessor) && (
            <div style={{ background: '#fff8e1', border: '1px solid #f0c36d', borderRadius: 6, padding: 12, marginBottom: 16 }}>
              <p style={{ marginTop: 0, fontWeight: 600 }}>Some processor information is missing</p>
              <p style={{ marginTop: 0, fontSize: 14 }}>Enter the processor for each item so it can be added to the first label line.</p>
              {parsedData.filter(item => item.needsProcessor).map((item) => (
                <label key={item.sourceKey} style={{ display: 'block', marginTop: 10 }}>
                  {item.topLine}
                  <input
                    type="text"
                    placeholder="Example: A17 Pro"
                    value={processorInputs[item.sourceKey] || ''}
                    onChange={(event) => updateProcessor(item.sourceKey, event.target.value)}
                    onBlur={(event) => confirmProcessor(item.sourceKey, event.target.value)}
                    style={{ display: 'block', width: '100%', boxSizing: 'border-box', padding: '8px 10px', marginTop: 4, border: '1px solid #cbd5e1', borderRadius: 6, fontSize: 15 }}
                  />
                </label>
              ))}
            </div>
          )}

          <label htmlFor="output-file-name" style={{ display: 'block', marginBottom: 6, fontWeight: 600 }}>
            File name
          </label>
          <input
            id="output-file-name"
            type="text"
            value={outputFileName}
            onChange={(event) => setOutputFileName(event.target.value)}
            placeholder="dymo_ready"
            style={{ width: '100%', boxSizing: 'border-box', padding: '10px 12px', marginBottom: 12, border: '1px solid #cbd5e1', borderRadius: 6, fontSize: 15 }}
          />

          <button
            onClick={handleExport}
            disabled={parsedData.some(item => item.needsProcessor)}
            style={{
              width: '100%',
              padding: '12px',
              backgroundColor: parsedData.some(item => item.needsProcessor) ? '#aab2bd' : '#34c759',
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
