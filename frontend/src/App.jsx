import React, { useState } from 'react';
import { UploadCloud, FileText, Wand2, Download, AlertCircle, CheckCircle2 } from 'lucide-react';

export default function App() {
  const [templateFile, setTemplateFile] = useState(null);
  const [articleFile, setArticleFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState('');

  const handleConvert = async () => {
    if (!templateFile || !articleFile) {
      setError('Please upload both a template and an article document');
      return;
    }
    setError('');
    setLoading(true);
    setSuccess(false);

    const formData = new FormData();
    formData.append('template_file', templateFile);
    formData.append('article_file', articleFile);

    try {
      const response = await fetch('http://localhost:8000/convert', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.detail || 'Failed to convert document');
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      setDownloadUrl(url);
      setSuccess(true);
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen relative overflow-hidden flex flex-col items-center justify-center p-6">
      {/* Background Orbs */}
      <div className="absolute top-1/4 left-1/4 w-96 h-96 bg-primary/20 rounded-full mix-blend-screen filter blur-[100px] anim-blob"></div>
      <div className="absolute top-1/3 right-1/4 w-96 h-96 bg-secondary/20 rounded-full mix-blend-screen filter blur-[100px] anim-blob animation-delay-2000"></div>
      <div className="absolute bottom-1/4 left-1/2 w-96 h-96 bg-blue-500/20 rounded-full mix-blend-screen filter blur-[100px] anim-blob animation-delay-4000"></div>

      <div className="z-10 w-full max-w-4xl glass-panel p-10 flex flex-col items-center">
        <div className="text-center space-y-4 mb-12 relative w-full">
          <h1 className="text-5xl font-extrabold tracking-tight bg-clip-text text-transparent bg-gradient-to-r from-white via-gray-300 to-gray-500">
            DocForge AI
          </h1>
          <p className="text-gray-400 text-lg max-w-2xl mx-auto">
            Intelligent semantic matching algorithm to map your article into any complex template architecture instantly.
          </p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-8 w-full mb-10">
          {/* Template Upload */}
          <label htmlFor="template-upload" className="cursor-pointer glass-panel p-6 flex flex-col items-center justify-center border border-white/10 hover:border-primary/50 transition-colors group relative overflow-hidden">
            <div className="pointer-events-none absolute inset-0 bg-gradient-to-br from-primary/5 to-transparent opacity-0 group-hover:opacity-100 transition-opacity"></div>
            <UploadCloud className="relative z-10 w-12 h-12 text-primary mb-4 group-hover:scale-110 transition-transform" />
            <h3 className="relative z-10 text-xl font-semibold mb-2">Target Template</h3>
            <p className="relative z-10 text-sm text-gray-400 text-center mb-6">Upload the .docx format you want to target</p>
            <input 
              type="file" 
              accept=".docx" 
              onChange={(e) => setTemplateFile(e.target.files[0])}
              className="hidden" 
              id="template-upload" 
            />
            <div className="relative z-10 border border-primary text-primary px-4 py-2 rounded-lg group-hover:bg-primary group-hover:text-white transition-colors text-center truncate max-w-full">
              {templateFile ? templateFile.name : 'Select Template'}
            </div>
          </label>

          {/* Source Article Upload */}
          <label htmlFor="article-upload" className="cursor-pointer glass-panel p-6 flex flex-col items-center justify-center border border-white/10 hover:border-secondary/50 transition-colors group relative overflow-hidden">
             <div className="pointer-events-none absolute inset-0 bg-gradient-to-bl from-secondary/5 to-transparent opacity-0 group-hover:opacity-100 transition-opacity"></div>
            <FileText className="relative z-10 w-12 h-12 text-secondary mb-4 group-hover:scale-110 transition-transform" />
            <h3 className="relative z-10 text-xl font-semibold mb-2">Source Article</h3>
            <p className="relative z-10 text-sm text-gray-400 text-center mb-6">Upload your raw manuscript in .docx</p>
            <input 
              type="file" 
              accept=".docx" 
              onChange={(e) => setArticleFile(e.target.files[0])}
              className="hidden" 
              id="article-upload" 
            />
            <div className="relative z-10 border border-secondary text-secondary px-4 py-2 rounded-lg group-hover:bg-secondary group-hover:text-white transition-colors text-center truncate max-w-full">
              {articleFile ? articleFile.name : 'Select Source'}
            </div>
          </label>
        </div>

        {error && (
          <div className="w-full mb-6 p-4 rounded-xl bg-red-500/10 border border-red-500/50 flex items-center text-red-200">
            <AlertCircle className="w-5 h-5 mr-3 shrink-0 text-red-400" />
            <span>{error}</span>
          </div>
        )}

        {success && (
          <div className="w-full mb-6 p-4 rounded-xl bg-green-500/10 border border-green-500/50 flex items-center justify-between text-green-200">
            <div className="flex items-center">
              <CheckCircle2 className="w-5 h-5 mr-3 shrink-0 text-green-400" />
              <span>Conversion successful! Your document is intelligently mapped.</span>
            </div>
            <a href={downloadUrl} download="Formatted_Document.docx" className="flex items-center space-x-2 bg-green-500/20 hover:bg-green-500/40 text-green-100 px-4 py-2 rounded-lg transition-colors">
              <Download className="w-4 h-4" />
              <span>Download</span>
            </a>
          </div>
        )}

        <button 
          onClick={handleConvert} 
          disabled={loading || success}
          className="btn-primary w-full max-w-sm flex items-center justify-center space-x-2 text-lg disabled:opacity-50"
        >
          {loading ? (
            <>
              <div className="w-5 h-5 border-2 border-white/30 border-t-white rounded-full animate-spin"></div>
              <span>Processing with AI...</span>
            </>
          ) : (
            <>
              <Wand2 className="w-5 h-5" />
              <span>Fuse Documents</span>
            </>
          )}
        </button>
      </div>

    </div>
  );
}
