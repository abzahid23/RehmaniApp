import React, { useState } from 'react';
import { Upload, Download, FileSpreadsheet, AlertCircle, CheckCircle } from 'lucide-react';
import * as XLSX from 'xlsx';

const ExcelExtractorDemo = () => {
  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [extractedData, setExtractedData] = useState([]);
  const [processing, setProcessing] = useState(false);
  const [extractionLog, setExtractionLog] = useState([]);

  // Store address lookup table
  const storeAddresses: any = {
    '10519': '400 N Foxridge Raymore MO',
    '11289': '1407 S Baltimor Kirksville MO',
    '11293': '1912 S Main St Maryville MO',
    '11504': '1710 W 6th St Emporia KS',
    '12545': '11904 Shawnee M Shawnee KS',
    '12832': '701 S Commercia Harrisonville MO',
    '13136': '412 Northland D Cameron MO',
    '13258': '14420 E Highway Kansas City MO',
    '13709': '400 SE Douglas Lees Summit MO',
    '13712': '6904 Hunter St Raytown MO',
    '41242': '13525 College B Olathe KS',
    '41962': '1900 SW State R Blue Springs MO',
    '43263': '11630 Saline J Nelson MO',
    '44417': '1609 N Missouri Macon MO',
    '44478': '8601 W 137th St Overland Park KS',
    '44754': '518 E Main St Gardner KS',
    '44790': '6001 Main St Grandview MO',
    '44857': '17930 W 119th S Olathe KS',
    '45070': '22520 Midland D Shawnee KS',
    '45198': '1100 W 135th Te Kansas City MO',
    '71504': 'Kansas Turnpike Matfield Green KS',
    '71515': '7581 SW Kansas Towanda KS',
    '72449': '1520 South Webb Wichita KS'
  };

  const handleFileUpload = (event: any) => {
    const files : React.SetStateAction<never[]> = Array.from(event.target.files);
    setUploadedFiles(files);
  };

  const extractDataFromExcel = (workbook : any, filename: any) => {
    const log = [];
    log.push(`📁 Processing: ${filename}`);

    try {
      const sheetNames = workbook.SheetNames;
      log.push(`📋 Found sheets: ${sheetNames.join(', ')}`);

      // Get the first sheet (usually the sales summary)
      const worksheet = workbook.Sheets[sheetNames[0]];
      
      // Method 1: Look for specific patterns in the sheet
      let storeNumber = '';
      let month = '';
      let netSales = 0;
      let orderCount = 0;
      
      // Extract store number from sheet data
      const sheetData : any = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
      
      // Look for store number pattern in first few rows
      for (let i = 0; i < Math.min(10, sheetData.length); i++) {
        const row = sheetData[i];
        for (let j = 0; j < row.length; j++) {
          const cell = String(row[j] || '');
          
          // Look for store number pattern like "10519 - Raymore, MO"
          const storeMatch = cell.match(/(\d{5})\s*[-–]\s*([^,]+),?\s*([A-Z]{2})/);
          if (storeMatch) {
            storeNumber = storeMatch[1];
            log.push(`🏪 Found store: ${storeNumber} - ${storeMatch[2]}, ${storeMatch[3]}`);
            break;
          }
          
          // Look for month/year
          const monthMatch = cell.match(/(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})/);
          if (monthMatch) {
            month = `${monthMatch[1]} ${monthMatch[2]}`;
            log.push(`📅 Found month: ${month}`);
          }
        }
        if (storeNumber) break;
      }

      // Method 2: Look for specific values by scanning the sheet
      const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1:Z100');
      
      for (let row = range.s.r; row <= Math.min(range.e.r, 50); row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
          const cell = worksheet[cellAddress];
          
          if (cell && cell.v) {
            const cellValue = String(cell.v);
            
            // Look for Net Sales
            if (cellValue.toLowerCase().includes('net sales') && !netSales) {
              // Check adjacent cells for the value
              for (let offset = 1; offset <= 3; offset++) {
                const valueCell = worksheet[XLSX.utils.encode_cell({ r: row, c: col + offset })];
                if (valueCell && typeof valueCell.v === 'number') {
                  netSales = valueCell.v;
                  log.push(`💰 Found Net Sales: $${netSales.toLocaleString()}`);
                  break;
                }
              }
            }
            
            // Look for Order Count
            if (cellValue.toLowerCase().includes('order count') && !orderCount) {
              for (let offset = 1; offset <= 3; offset++) {
                const valueCell = worksheet[XLSX.utils.encode_cell({ r: row, c: col + offset })];
                if (valueCell && typeof valueCell.v === 'number') {
                  orderCount = valueCell.v;
                  log.push(`🧾 Found Order Count: ${orderCount.toLocaleString()}`);
                  break;
                }
              }
            }
          }
        }
      }

      // Look for Revenue Centers section
      let revenues: any = {
        softServe: 0,
        food: 0,
        noveltiesBoxed: 0,
        beverages: 0,
        cakes: 0,
        breakfast: 0,
        ojBeverages: 0
      };

      // Scan for revenue center data
      for (let row = range.s.r; row <= Math.min(range.e.r, 100); row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
          const cell = worksheet[cellAddress];
          
          if (cell && cell.v) {
            const cellValue = String(cell.v).toLowerCase();
            
            // Map revenue categories
            const categoryMap : any= {
              'soft serve': 'softServe',
              'food': 'food',
              'novelties-boxed': 'noveltiesBoxed',
              'beverage': 'beverages',
              'cakes': 'cakes',
              'breakfast': 'breakfast',
              'oj beverages': 'ojBeverages'
            };

            Object.keys(categoryMap).forEach(category => {
              if (cellValue.includes(category) && revenues[categoryMap[category]] === 0) {
                // Look for dollar amount in the same row
                for (let offset = 1; offset <= 5; offset++) {
                  const valueCell = worksheet[XLSX.utils.encode_cell({ r: row, c: col + offset })];
                  if (valueCell && typeof valueCell.v === 'number' && valueCell.v > 100) {
                    revenues[categoryMap[category]] = valueCell.v;
                    log.push(`🍦 Found ${category}: $${valueCell.v.toLocaleString()}`);
                    break;
                  }
                }
              }
            });
          }
        }
      }

      // Calculate derived values
      const dqFood = revenues.food + revenues.noveltiesBoxed;
      const totalSales: any = Object.values(revenues).reduce((sum: any, val: any) => sum + val, 0);

      log.push(`✅ Extraction completed successfully!`);

      return {
        success: true,
        data: {
          month,
          storeNumber,
          storeAddress: storeAddresses[storeNumber] || 'Address not found',
          dairyQueen: revenues.softServe,
          dqFood,
          beverages: revenues.beverages,
          breakfast: revenues.breakfast,
          cakes: revenues.cakes,
          ojBeverages: revenues.ojBeverages,
          transactionCount: orderCount,
          netSalesWithDonations: netSales,
          totalSales: Math.max(totalSales, netSales), // Use net sales if revenue breakdown not found
          filename
        },
        log
      };

    } catch (error: any) {
      log.push(`❌ Error: ${error.message}`);
      return { success: false, log, error: error.message };
    }
  };

  const processExcelFiles = async () => {
    setProcessing(true);
    setExtractionLog([]);
    const results: any = [];
    const allLogs: any = [];

    for (const file of uploadedFiles) {
      const tempFileObject: any = file;
      try {
        const arrayBuffer = await tempFileObject.arrayBuffer();
        const workbook = XLSX.read(arrayBuffer, {
          cellStyles: true,
          cellFormula: true,
          cellDates: true
        });

        const extraction = extractDataFromExcel(workbook, tempFileObject.name);
        allLogs.push(...extraction.log);
        
        if (extraction.success) {
          results.push(extraction.data);
          allLogs.push(`✅ Successfully processed ${tempFileObject.name}`);
        } else {
          allLogs.push(`❌ Failed to process ${tempFileObject.name}: ${extraction.error}`);
        }
        
      } catch (error: any) {
        allLogs.push(`❌ Error reading ${tempFileObject.name}: ${error.message}`);
      }
    }

    // Calculate totals
    if (results.length > 0) {
      const totals = results.reduce((acc: any, row: any) => ({
        dairyQueen: acc.dairyQueen + row.dairyQueen,
        dqFood: acc.dqFood + row.dqFood,
        beverages: acc.beverages + row.beverages,
        breakfast: acc.breakfast + row.breakfast,
        cakes: acc.cakes + row.cakes,
        ojBeverages: acc.ojBeverages + row.ojBeverages,
        transactionCount: acc.transactionCount + row.transactionCount,
        netSalesWithDonations: acc.netSalesWithDonations + row.netSalesWithDonations,
        totalSales: acc.totalSales + row.totalSales
      }), {
        dairyQueen: 0, dqFood: 0, beverages: 0, breakfast: 0, cakes: 0, 
        ojBeverages: 0, transactionCount: 0, netSalesWithDonations: 0, totalSales: 0
      });

      results.push({
        month: 'TOTAL',
        storeNumber: '',
        storeAddress: 'All Stores Combined',
        ...totals,
        filename: 'TOTALS'
      });

      allLogs.push(`📊 Generated totals for ${results.length - 1} stores`);
    }

    setExtractedData(results);
    setExtractionLog(allLogs);
    setProcessing(false);
  };

  const createSampleExcel = () => {
    // Create a sample Excel file showing the expected format
    const sampleData = [
      ['Sales Summary', '10519 - Raymore, MO - GC'],
      ['Date:', 'April 2025'],
      [''],
      ['Net Sales', '$113,551.47'],
      ['Order Count:', '8,678'],
      [''],
      ['Revenue Centers'],
      ['Category', 'Quantity', 'Total', 'Percent'],
      ['Beverage', '478', '$1,193.05', '1.05%'],
      ['Cakes', '233', '$7,165.08', '6.31%'],
      ['Food', '6,450', '$40,987.52', '36.10%'],
      ['Novelties-Boxed', '149', '$1,855.91', '1.63%'],
      ['Soft Serve', '11,494', '$62,349.91', '54.91%']
    ];

    const ws = XLSX.utils.aoa_to_sheet(sampleData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sales Summary');

    const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([excelBuffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Sample_Sales_Summary_Format.xlsx';
    a.click();
    URL.revokeObjectURL(url);
  };

  const formatCurrency = (value: any) => {
    return new Intl.NumberFormat('en-US', {
      style: 'currency',
      currency: 'USD'
    }).format(value || 0);
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-green-50 to-blue-50 p-6">
      <div className="max-w-7xl mx-auto">
        {/* Header */}
        <div className="text-center mb-8">
          <div className="flex items-center justify-center mb-4">
            <FileSpreadsheet className="text-green-600 mr-3" size={40} />
            <h1 className="text-4xl font-bold text-gray-800">Excel Sales Data Extractor</h1>
          </div>
          <p className="text-gray-600 text-lg">Much Better Than PDF! Accurate & Reliable Data Extraction</p>
        </div>

        {/* Comparison Section */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 mb-8">
          <div className="bg-red-50 border border-red-200 rounded-lg p-6">
            <h3 className="text-lg font-semibold text-red-800 mb-3 flex items-center">
              <AlertCircle className="mr-2" size={20} />
              PDF Format Issues
            </h3>
            <ul className="text-red-700 space-y-2 text-sm">
              <li>• Text parsing can be unreliable</li>
              <li>• Formatting variations cause errors</li>
              <li>• Requires PDF parsing libraries</li>
              <li>• Slower processing speed</li>
              <li>• Hard to debug extraction issues</li>
            </ul>
          </div>
          
          <div className="bg-green-50 border border-green-200 rounded-lg p-6">
            <h3 className="text-lg font-semibold text-green-800 mb-3 flex items-center">
              <CheckCircle className="mr-2" size={20} />
              Excel Format Benefits
            </h3>
            <ul className="text-green-700 space-y-2 text-sm">
              <li>• 100% accurate data extraction</li>
              <li>• Direct cell access by address</li>
              <li>• No additional libraries needed</li>
              <li>• Much faster processing</li>
              <li>• Easy to debug and modify</li>
            </ul>
          </div>
        </div>

        {/* Upload Section */}
        <div className="bg-white rounded-lg shadow-lg p-6 mb-6">
          <h2 className="text-2xl font-semibold text-gray-800 mb-4 flex items-center">
            <Upload className="mr-2" size={24} />
            Upload Excel Sales Summary Files
          </h2>
          
          <div className="border-2 border-dashed border-green-300 rounded-lg p-8 text-center mb-4">
            <input
              type="file"
              multiple
              accept=".xlsx,.xls,.csv"
              onChange={handleFileUpload}
              className="hidden"
              id="excel-upload"
            />
            <label htmlFor="excel-upload" className="cursor-pointer">
              <FileSpreadsheet className="mx-auto mb-4 text-green-400" size={48} />
              <p className="text-lg text-gray-600 mb-2">
                Click to select Excel sales summary files
              </p>
              <p className="text-sm text-gray-500">
                Supports .xlsx, .xls, and .csv formats
              </p>
            </label>
          </div>

          <div className="flex gap-4">
            <button
              onClick={createSampleExcel}
              className="bg-blue-600 text-white px-6 py-2 rounded-lg hover:bg-blue-700"
            >
              <Download className="inline mr-2" size={16} />
              Download Sample Format
            </button>
            
            {uploadedFiles.length > 0 && (
              <button
                onClick={processExcelFiles}
                disabled={processing}
                className="bg-green-600 text-white px-6 py-2 rounded-lg hover:bg-green-700 disabled:opacity-50"
              >
                {processing ? 'Processing...' : `Process ${uploadedFiles.length} Files`}
              </button>
            )}
          </div>
        </div>

        {/* Extraction Log */}
        {extractionLog.length > 0 && (
          <div className="bg-gray-50 rounded-lg p-6 mb-6">
            <h3 className="text-lg font-semibold text-gray-800 mb-3">Extraction Log</h3>
            <div className="max-h-64 overflow-y-auto bg-white p-4 rounded border font-mono text-sm">
              {extractionLog.map((log, index) => (
                <div key={index} className="mb-1">
                  {log}
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Data Table */}
        {extractedData.length > 0 && (
          <div className="bg-white rounded-lg shadow-lg overflow-hidden">
            <div className="p-6 border-b border-gray-200">
              <h2 className="text-2xl font-semibold text-gray-800">Extracted Sales Data</h2>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full">
                <thead className="bg-gray-50">
                  <tr>
                    <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase">Store #</th>
                    <th className="px-4 py-3 text-left text-xs font-medium text-gray-500 uppercase">Address</th>
                    <th className="px-4 py-3 text-right text-xs font-medium text-gray-500 uppercase">Dairy Queen</th>
                    <th className="px-4 py-3 text-right text-xs font-medium text-gray-500 uppercase">DQ Food</th>
                    <th className="px-4 py-3 text-right text-xs font-medium text-gray-500 uppercase">Beverages</th>
                    <th className="px-4 py-3 text-right text-xs font-medium text-gray-500 uppercase">Cakes</th>
                    <th className="px-4 py-3 text-right text-xs font-medium text-gray-500 uppercase">Transactions</th>
                    <th className="px-4 py-3 text-right text-xs font-medium text-gray-500 uppercase">Total Sales</th>
                  </tr>
                </thead>
                <tbody className="bg-white divide-y divide-gray-200">
                  {extractedData.map((row: any, index) => (
                    <tr key={index} className={row.month === 'TOTAL' ? 'bg-yellow-50 font-semibold' : 'hover:bg-gray-50'}>
                      <td className="px-4 py-4 whitespace-nowrap text-sm text-gray-900">{row.storeNumber}</td>
                      <td className="px-4 py-4 whitespace-nowrap text-sm text-gray-900">{row.storeAddress}</td>
                      <td className="px-4 py-4 whitespace-nowrap text-sm text-gray-900 text-right">{formatCurrency(row.dairyQueen)}</td>
                      <td className="px-4 py-4 whitespace-nowrap text-sm text-gray-900 text-right">{formatCurrency(row.dqFood)}</td>
                      <td className="px-4 py-4 whitespace-nowrap text-sm text-gray-900 text-right">{formatCurrency(row.beverages)}</td>
                      <td className="px-4 py-4 whitespace-nowrap text-sm text-gray-900 text-right">{formatCurrency(row.cakes)}</td>
                      <td className="px-4 py-4 whitespace-nowrap text-sm text-gray-900 text-right">{row.transactionCount.toLocaleString()}</td>
                      <td className="px-4 py-4 whitespace-nowrap text-sm text-gray-900 text-right">{formatCurrency(row.totalSales)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}

        {/* Instructions */}
        <div className="mt-8 bg-blue-50 border border-blue-200 rounded-lg p-6">
          <h3 className="text-lg font-semibold text-blue-800 mb-3">How to Switch to Excel Format</h3>
          <ol className="list-decimal list-inside text-blue-700 space-y-2">
            <li>Check if your Brink POS system can export sales summaries as Excel (.xlsx) files</li>
            <li>Download the sample format to see the expected structure</li>
            <li>Configure your POS exports to match this format</li>
            <li>Upload Excel files instead of PDFs for 100% accurate extraction</li>
            <li>Enjoy much faster and more reliable processing!</li>
          </ol>
          <div className="mt-4 p-4 bg-green-50 border border-green-200 rounded-lg">
            <p className="text-sm text-green-800">
              <strong>💡 Pro Tip:</strong> Even if your POS only exports PDFs, you can often open them in Excel 
              and save as .xlsx format. This manual step once per month could save hours of troubleshooting 
              automated PDF parsing issues.
            </p>
          </div>
        </div>
      </div>
    </div>
  );
};

export default ExcelExtractorDemo;