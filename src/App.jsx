import React, { useState, useRef } from 'react';
import { UploadCloud, FileSpreadsheet, Download, CheckCircle2 } from 'lucide-react';
import * as XLSX from 'xlsx';
import './App.css';

function App() {
  const [file, setFile] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const [isProcessing, setIsProcessing] = useState(false);
  const [isDone, setIsDone] = useState(false);
  const fileInputRef = useRef(null);

  const handleDragOver = (e) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      setFile(e.dataTransfer.files[0]);
      setIsDone(false);
    }
  };

  const handleFileSelect = (e) => {
    if (e.target.files && e.target.files.length > 0) {
      setFile(e.target.files[0]);
      setIsDone(false);
    }
  };

  const processFile = async () => {
    if (!file) return;
    setIsProcessing(true);
    
    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheetName];
      const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
      
      const outputData = json.map(row => {
        const getVal = (key) => row[key] || "";
        
        const pan = getVal("PAN Number");
        const consultantEmployeeType = getVal("Consultant / Employee").toString().toLowerCase().trim();
        const salaryText = getVal("Salary") || getVal("Amount") || "0"; 
        // Note: Using a heuristic since exact column wasn't in original list, assumed to be part of the dummy sheet
        const amountOfPayment = parseFloat(salaryText.replace(/[^0-9.]/g, '')) || 0;
        
        let section = "";
        let challanPlaceholder = "";
        
        if (consultantEmployeeType.includes("consultant")) {
          section = "194J";
          challanPlaceholder = "A";
        } else if (consultantEmployeeType.includes("employee")) {
          section = "192B";
          challanPlaceholder = "B";
        } else if (consultantEmployeeType.includes("vendor")) {
          section = "194C"; // Defaulting to 194C (could also be 194I depending on other logic)
          challanPlaceholder = "C";
        }
        
        // Vendor Taxpayer type check (4th character of PAN)
        // C for Company/Corporate, anything else O for other
        let deducteeCode = "O";
        if (pan.length >= 4) {
          const fourthChar = pan.charAt(3).toUpperCase();
          if (fourthChar === 'C') {
            deducteeCode = "C";
          }
        }
        
        // Surcharge calculation: 1.5% of the payment amount
        const surcharge = (amountOfPayment * 0.015).toFixed(2);
        
        return {
          "Section": section, 
          "New Code": "",
          "Deductee Code (\"C\" for Corporate\", \"O\" for Other\")": deducteeCode,
          "Permanent Account Number (PAN) of deductee": pan,
          "Name of Deductee": getVal("Name"),
          "FC": "", // FCRA / Non-FCRA to be provided by Finance team
          "Deductee  Address1": "",
          "Deductee  Address2": "",
          "Deductee  Address3": "",
          "Deductee  Address4": "",
          "Deductee  City": "",
          "Deductee  State": "",
          "Deductee  PIN": "",
          "Amount of Payment  ( Rs.)": amountOfPayment || "",
          "Date on which Amount paid / credited": "", // To be filled as last day of the month by finance team
          "Rate at which tax deducted": "",
          "TDS (Rs.)": "",
          "Surcharge (Rs.)": surcharge,
          "Education Cess (Rs.)": "",
          "Date on which tax deducted": "",
          "Date on which tax paid to the credit of Central Govt": "", // Finance team info
          "Bank-Branch Code": "",
          "Challan No given by Bank / Tr. Voucher No.": challanPlaceholder,
          "Date of furnishing Tax Deduction Certificate": ""
        };
      });
      
      const newWorksheet = XLSX.utils.json_to_sheet(outputData);
      const newWorkbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Output");
      XLSX.writeFile(newWorkbook, "TDS_Output.xlsx");
      
      setIsDone(true);
    } catch (error) {
      console.error("Error processing file", error);
      alert("There was an error processing the file. Please ensure it's a valid Excel/CSV file.");
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <div className="app-container">
      <header className="header">
        <h1>Data Extractor</h1>
        <p>Automate your TDS data preparation securely in the browser.</p>
      </header>

      <main className="upload-card">
        {!file ? (
          <div 
            className={`upload-area ${isDragging ? 'dragging' : ''}`}
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
          >
            <UploadCloud size={48} className="upload-icon" />
            <h3 className="upload-text">Click or drag Excel file here</h3>
            <p className="upload-subtext">Supports .xlsx, .xls, .csv</p>
            <input 
              type="file" 
              className="file-input" 
              ref={fileInputRef} 
              accept=".csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
              onChange={handleFileSelect}
            />
          </div>
        ) : (
          <div className="process-section">
            <div className="file-info">
              <FileSpreadsheet size={24} className="upload-icon" />
              <span>{file.name}</span>
            </div>
            
            {isDone ? (
              <>
                <div style={{ color: '#10B981', display: 'flex', alignItems: 'center', gap: '8px', margin: '1rem 0' }}>
                  <CheckCircle2 size={24} />
                  <span>Processing complete! File downloaded.</span>
                </div>
                <button className="btn-primary" onClick={() => { setFile(null); setIsDone(false); }}>
                  Upload Another File
                </button>
              </>
            ) : (
              <button 
                className="btn-primary" 
                onClick={processFile}
                disabled={isProcessing}
              >
                {isProcessing ? 'Processing...' : (
                  <>
                    <Download size={20} />
                    Extract Data
                  </>
                )}
              </button>
            )}
          </div>
        )}
      </main>
    </div>
  );
}

export default App;
