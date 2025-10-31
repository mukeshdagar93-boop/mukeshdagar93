
import React, { useState, useCallback, useEffect, useRef } from 'react';
import { ManagedFile } from './types';
import { mergePdfs } from './services/pdfService';
import FileUpload from './components/FileUpload';
import FileList from './components/FileList';
import { SpinnerIcon, DownloadIcon, PlusIcon } from './components/Icons';

const App: React.FC = () => {
  const [files, setFiles] = useState<ManagedFile[]>([]);
  const [isMerging, setIsMerging] = useState<boolean>(false);
  const [mergedPdfUrl, setMergedPdfUrl] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    // Clean up the object URL to avoid memory leaks
    return () => {
      if (mergedPdfUrl) {
        URL.revokeObjectURL(mergedPdfUrl);
      }
    };
  }, [mergedPdfUrl]);

  const handleFilesAdded = useCallback((newFiles: File[]) => {
    setError(null);
    const pdfFiles = newFiles
      .filter(file => file.type === 'application/pdf')
      .map(file => ({
        id: `${file.name}-${file.lastModified}-${Math.random()}`,
        file,
      }));
    
    if (pdfFiles.length !== newFiles.length) {
      setError('Some files were not PDFs and were ignored.');
    }

    setFiles(prev => [...prev, ...pdfFiles]);
  }, []);

  const handleRemoveFile = useCallback((id: string) => {
    setFiles(prev => prev.filter(f => f.id !== id));
  }, []);

  const handleReorderFiles = useCallback((reorderedFiles: ManagedFile[]) => {
    setFiles(reorderedFiles);
  }, []);

  const handleMerge = async () => {
    if (files.length < 2) {
      setError('Please upload at least two PDF files to merge.');
      return;
    }
    setIsMerging(true);
    setError(null);
    if (mergedPdfUrl) {
      URL.revokeObjectURL(mergedPdfUrl);
      setMergedPdfUrl(null);
    }

    try {
      const fileArray = files.map(f => f.file);
      const mergedPdfBytes = await mergePdfs(fileArray);
      const blob = new Blob([mergedPdfBytes], { type: 'application/pdf' });
      const url = URL.createObjectURL(blob);
      setMergedPdfUrl(url);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An unknown error occurred during merging.');
      console.error(err);
    } finally {
      setIsMerging(false);
    }
  };

  const handleStartOver = () => {
    setFiles([]);
    if (mergedPdfUrl) {
      URL.revokeObjectURL(mergedPdfUrl);
    }
    setMergedPdfUrl(null);
    setError(null);
    setIsMerging(false);
  };
  
  const handleAddMoreFilesClick = () => {
    fileInputRef.current?.click();
  };

  const renderContent = () => {
    if (isMerging) {
      return (
        <div className="text-center p-10">
          <SpinnerIcon className="w-16 h-16 mx-auto text-brand-primary" />
          <h2 className="mt-4 text-2xl font-semibold">Merging PDFs...</h2>
          <p className="text-slate-400">Please wait while we combine your documents.</p>
        </div>
      );
    }

    if (mergedPdfUrl) {
      return (
        <div className="text-center p-10 bg-slate-800 rounded-lg">
          <h2 className="text-3xl font-bold text-green-400">Merge Successful!</h2>
          <p className="mt-2 text-slate-300">Your combined PDF is ready for download.</p>
          <div className="mt-8 flex flex-col sm:flex-row justify-center gap-4">
            <a
              href={mergedPdfUrl}
              download="merged.pdf"
              className="inline-flex items-center justify-center px-8 py-3 bg-brand-primary text-white font-bold rounded-lg hover:bg-brand-secondary transition-colors text-lg"
            >
              <DownloadIcon className="mr-3" />
              Download PDF
            </a>
            <button
              onClick={handleStartOver}
              className="px-8 py-3 bg-slate-600 text-white font-bold rounded-lg hover:bg-slate-700 transition-colors"
            >
              Start Over
            </button>
          </div>
        </div>
      );
    }

    if (files.length === 0) {
      return (
        <div>
          <h1 className="text-4xl md:text-5xl font-extrabold text-center mb-2">PDF for Kamlesh Mehta</h1>
          <p className="text-lg text-slate-400 text-center mb-8">Combine multiple PDFs into one, effortlessly.</p>
          <FileUpload onFilesAdded={handleFilesAdded} />
        </div>
      );
    }

    return (
      <div>
        <h1 className="text-3xl font-bold mb-6">Your Files</h1>
        <p className="text-slate-400 mb-4">Drag and drop files to change the merge order.</p>
        <div className="mb-6">
            <FileList files={files} onReorder={handleReorderFiles} onRemove={handleRemoveFile} />
        </div>
        <div className="flex flex-col sm:flex-row gap-4">
             <input
                ref={fileInputRef}
                type="file"
                multiple
                accept=".pdf"
                className="hidden"
                onChange={(e) => handleFilesAdded(Array.from(e.target.files || []))}
              />
            <button
                onClick={handleAddMoreFilesClick}
                className="flex-1 inline-flex items-center justify-center px-6 py-3 bg-slate-700 text-white font-semibold rounded-lg hover:bg-slate-600 transition-colors"
            >
              <PlusIcon className="mr-2" />
              Add More Files
            </button>
            <button
              onClick={handleMerge}
              disabled={files.length < 2}
              className="flex-1 inline-flex items-center justify-center px-6 py-4 bg-brand-primary text-white font-bold rounded-lg hover:bg-brand-secondary transition-colors disabled:bg-slate-700 disabled:cursor-not-allowed disabled:text-slate-400 text-lg"
            >
              Merge {files.length} {files.length === 1 ? 'File' : 'Files'}
            </button>
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen p-4 sm:p-8 flex items-center justify-center bg-dots">
      <style>{`.bg-dots { background-image: radial-gradient(circle, rgba(255,255,255,0.05) 1px, transparent 1px); background-size: 20px 20px; }`}</style>
      <main className="w-full max-w-3xl mx-auto bg-slate-900/70 backdrop-blur-sm border border-slate-700 rounded-xl shadow-2xl p-6 sm:p-8">
        {error && (
            <div className="bg-red-500/20 border border-red-500 text-red-300 px-4 py-3 rounded-lg relative mb-6" role="alert">
                <strong className="font-bold">Error: </strong>
                <span className="block sm:inline">{error}</span>
                <button onClick={() => setError(null)} className="absolute top-0 bottom-0 right-0 px-4 py-3">
                    <span className="text-2xl">&times;</span>
                </button>
            </div>
        )}
        {renderContent()}
      </main>
    </div>
  );
};

export default App;
