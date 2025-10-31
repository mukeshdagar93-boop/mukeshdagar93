
import React, { useState, useCallback } from 'react';
import FileUpload from './components/FileUpload';
import Loader from './components/Loader';
import { ExcelIcon } from './components/icons/ExcelIcon';
import { PdfIcon } from './components/icons/PdfIcon';
import { CheckCircleIcon } from './components/icons/CheckCircleIcon';
import { XCircleIcon } from './components/icons/XCircleIcon';

// Add declarations for CDN-loaded libraries to avoid TypeScript errors.
declare global {
  interface Window {
    XLSX: any;
    jspdf: any;
  }
}

type Status = 'idle' | 'file-selected' | 'converting' | 'success' | 'error';

const App: React.FC = () => {
  const [file, setFile] = useState<File | null>(null);
  const [status, setStatus] = useState<Status>('idle');
  const [errorMessage, setErrorMessage] = useState<string>('');
  const [pdfBlobUrl, setPdfBlobUrl] = useState<string | null>(null);

  const handleFileChange = (selectedFile: File | null) => {
    if (selectedFile) {
      const allowedMimeTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', // .xlsx
        'application/vnd.ms-excel', // .xls
        'application/vnd.ms-excel.sheet.macroEnabled.12', // .xlsm
        'application/vnd.ms-excel.binary.macroEnabled.12', // .xlsb
        'text/csv',
        'application/vnd.oasis.opendocument.spreadsheet', // .ods
      ];
      const allowedExtensions = ['.xlsx', '.xls', '.xlsm', '.xlsb', '.csv', '.ods'];
      const fileName = selectedFile.name.toLowerCase();
      const fileIsValid = allowedMimeTypes.includes(selectedFile.type) ||
        allowedExtensions.some(ext => fileName.endsWith(ext));

      if (fileIsValid) {
        setFile(selectedFile);
        setStatus('file-selected');
        setErrorMessage('');
        setPdfBlobUrl(null);
      } else {
        setErrorMessage('Invalid file type. Please upload an Excel, CSV, or ODS file.');
        setStatus('error');
        setFile(null);
      }
    }
  };

  const handleReset = useCallback(() => {
    setFile(null);
    setStatus('idle');
    setErrorMessage('');
    if (pdfBlobUrl) {
      URL.revokeObjectURL(pdfBlobUrl);
    }
    setPdfBlobUrl(null);
  }, [pdfBlobUrl]);

  const handleConvert = useCallback(async () => {
    if (!file) return;

    setStatus('converting');
    try {
      const reader = new FileReader();
      reader.readAsArrayBuffer(file);
      reader.onload = (e) => {
        if (!e.target?.result) {
          setErrorMessage('Failed to read the file.');
          setStatus('error');
          return;
        }

        const data = new Uint8Array(e.target.result as ArrayBuffer);
        // Add cellStyles: true to parse styles from the Excel file
        const workbook = window.XLSX.read(data, { type: 'array', cellStyles: true });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Use `raw: false` to get formatted text for dates, currency, etc.
        const json_data: any[][] = window.XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false });
        
        if (json_data.length === 0) {
            setErrorMessage('The selected Excel sheet is empty.');
            setStatus('error');
            return;
        }

        const doc = new window.jspdf.jsPDF({
            orientation: 'landscape',
        });

        const tableHead = json_data[0] || [];
        const tableBody = json_data.length > 1 ? json_data.slice(1) : [];

        (doc as any).autoTable({
          head: [tableHead],
          body: tableBody,
          startY: 10,
          theme: 'grid',
          tableWidth: 'auto', // For auto-fitting columns
          styles: {
            fontSize: 8,
            cellPadding: 2,
            overflow: 'linebreak', // To wrap text and auto-fit columns
          },
          headStyles: {
            fillColor: [22, 163, 74], // green-600
            textColor: 255,
            fontStyle: 'bold',
          },
          didParseCell: (data: any) => {
            const rowIndex = data.section === 'head' ? data.row.index : data.row.index + 1;
            const colIndex = data.column.index;
            
            const cellAddress = window.XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
            const excelCell = worksheet[cellAddress];
        
            if (excelCell && excelCell.s) {
                const styles = excelCell.s;
                
                // Font style
                if (styles.font) {
                    let fontStyle = '';
                    if (styles.font.bold) fontStyle += 'bold';
                    if (styles.font.italic) fontStyle += 'italic';
                    if (fontStyle) data.cell.styles.fontStyle = fontStyle;
                    
                    if (styles.font.color?.rgb) {
                        const hexColor = `#${styles.font.color.rgb.slice(-6)}`;
                        data.cell.styles.textColor = hexColor;
                    }
                }
                
                // Fill color
                if (styles.fill?.fgColor?.rgb) {
                    const rgb = styles.fill.fgColor.rgb;
                    // Ignore theme colors and default/white fills
                    if (rgb && !styles.fill.fgColor.theme && rgb !== '00000000' && rgb !== 'FFFFFFFF') {
                        const hexColor = `#${rgb.slice(-6)}`;
                        data.cell.styles.fillColor = hexColor;
                    }
                }
                
                // Alignment
                if (styles.alignment) {
                    if (styles.alignment.horizontal) {
                        data.cell.styles.halign = styles.alignment.horizontal;
                    }
                    if (styles.alignment.vertical) {
                        data.cell.styles.valign = styles.alignment.vertical;
                    }
                }
            }
          }
        });

        const url = doc.output('bloburl');
        setPdfBlobUrl(url.toString());
        setStatus('success');
      };
    } catch (err) {
      console.error(err);
      setErrorMessage('An unexpected error occurred during conversion.');
      setStatus('error');
    }
  }, [file]);

  const formatFileSize = (bytes: number): string => {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  };

  const renderContent = () => {
    switch (status) {
      case 'idle':
        return <FileUpload onFileSelect={handleFileChange} />;
      case 'file-selected':
      case 'converting':
      case 'success':
      case 'error':
        return (
          <div className="w-full max-w-lg bg-white dark:bg-slate-800 rounded-2xl shadow-xl p-8 text-center transition-all duration-300 ease-in-out">
            {status !== 'idle' && file && (
               <div className="flex items-center justify-between text-slate-600 dark:text-slate-300 border-b border-slate-200 dark:border-slate-700 pb-4 mb-4">
                 <div className="flex items-center space-x-3 text-left">
                  <ExcelIcon className="h-10 w-10 text-green-500" />
                  <div>
                    <p className="font-semibold truncate max-w-[200px] sm:max-w-xs">{file.name}</p>
                    <p className="text-sm text-slate-500 dark:text-slate-400">{formatFileSize(file.size)}</p>
                  </div>
                 </div>
                 <button onClick={handleReset} className="p-1 text-slate-400 hover:text-red-500 dark:hover:text-red-400 rounded-full focus:outline-none focus:ring-2 focus:ring-red-500 focus:ring-offset-2 dark:focus:ring-offset-slate-800">
                    <XCircleIcon className="h-7 w-7"/>
                    <span className="sr-only">Remove file</span>
                </button>
               </div>
            )}
            
            {status === 'converting' && (
              <div className="flex flex-col items-center justify-center space-y-4 h-40">
                <Loader />
                <p className="text-slate-500 dark:text-slate-400 font-medium">Converting your file...</p>
              </div>
            )}
            
            {status === 'success' && pdfBlobUrl && (
              <div className="flex flex-col items-center justify-center space-y-5 h-40">
                <CheckCircleIcon className="h-16 w-16 text-green-500"/>
                <p className="text-xl font-semibold text-slate-800 dark:text-slate-100">Conversion Successful!</p>
                <a
                  href={pdfBlobUrl}
                  download={`${file?.name.split('.')[0] || 'converted'}.pdf`}
                  className="w-full flex items-center justify-center space-x-2 bg-green-600 text-white font-bold py-3 px-4 rounded-lg hover:bg-green-700 focus:outline-none focus:ring-4 focus:ring-green-300 dark:focus:ring-green-800 transition-transform transform hover:scale-105"
                >
                  <PdfIcon className="h-6 w-6"/>
                  <span>Download PDF</span>
                </a>
              </div>
            )}

            {(status === 'file-selected') && (
              <div className="flex flex-col items-center justify-center space-y-4 h-40">
                <p className="text-slate-500 dark:text-slate-400">Your file is ready to be converted.</p>
                <button
                  onClick={handleConvert}
                  disabled={status === 'converting'}
                  className="w-full bg-indigo-600 text-white font-bold py-3 px-4 rounded-lg hover:bg-indigo-700 focus:outline-none focus:ring-4 focus:ring-indigo-300 dark:focus:ring-indigo-800 disabled:bg-indigo-300 disabled:cursor-not-allowed transition-transform transform hover:scale-105"
                >
                  Convert to PDF
                </button>
              </div>
            )}
            
            {status === 'error' && (
                <div className="flex flex-col items-center justify-center space-y-4 h-40 text-red-500 dark:text-red-400 bg-red-50 dark:bg-red-900/20 rounded-lg p-4">
                    <XCircleIcon className="h-12 w-12"/>
                    <p className="font-semibold">Conversion Failed</p>
                    <p className="text-sm">{errorMessage}</p>
                </div>
            )}

          </div>
        );
      default:
        return null;
    }
  };

  return (
    <div className="min-h-screen flex flex-col items-center justify-center p-4">
      <header className="text-center mb-8">
        <h1 className="text-4xl md:text-5xl font-extrabold text-slate-900 dark:text-white">
          Excel to PDF Converter
        </h1>
        <p className="mt-4 text-lg text-slate-600 dark:text-slate-300 max-w-2xl">
          Instantly convert your Excel spreadsheets into high-quality PDF documents online, for free.
        </p>
      </header>
      <main className="w-full flex items-center justify-center">
        {renderContent()}
      </main>
      <footer className="text-center mt-8 text-slate-500 dark:text-slate-400 text-sm">
        <p>&copy; {new Date().getFullYear()} ExcelToPDF. All rights reserved.</p>
      </footer>
    </div>
  );
};

export default App;
