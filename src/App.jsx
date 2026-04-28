import { useState } from 'react';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';
import './App.css';

// Simple Icon Components
const Icons = {
  Bank: () => <span>🏦</span>,
  Users: () => <span>👥</span>,
  Download: () => <span>⬇️</span>,
  Upload: () => <span>📤</span>,
  Check: () => <span>✓</span>,
  Info: () => <span>ⓘ</span>,
  Sparkles: () => <span>✨</span>,
  ArrowLeft: () => <span>←</span>,
  File: () => <span>📁</span>,
};

function App() {
  const [templateFile, setTemplateFile] = useState(null);
  const [excelNames, setExcelNames] = useState([]);
  const [manualNames, setManualNames] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [hasGenerated, setHasGenerated] = useState(false);
  const [downloadData, setDownloadData] = useState({ blob: null, fileName: '', isZip: false });
  const [activeStep, setActiveStep] = useState(1);

  // State disesuaikan dengan placeholder di Surat Konfirmasi Bank.docx
  const [formData, setFormData] = useState({
    Tanggal_Konfirmasi: '', 
    Nama_Klien: '',
    sebutan1: '', 
    Auditor1: '', 
    sebutan2: '', 
    Auditor2: '',
    Nama_Direktur: '', 
    Jabatan: ''
  });

  const handleInputChange = (e) => {
    const { name, value } = e.target;
    setFormData((prev) => ({ ...prev, [name]: value }));
  };

  const handleExcelUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheet = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[firstSheet];
      const json = XLSX.utils.sheet_to_json(worksheet);
      const names = [];
      json.forEach((row) => {
        const nameKey = Object.keys(row).find((key) => 
          key.toLowerCase().includes('nama') || key.toLowerCase().includes('bank')
        );
        if (nameKey && row[nameKey]) names.push(row[nameKey].toString().trim());
      });
      setExcelNames(names);
      setActiveStep(3);
      alert(`${names.length} nama bank berhasil diimport!`);
    };
    reader.readAsArrayBuffer(file);
  };

  const generateDocuments = async () => {
    if (!templateFile) {
      alert('Harap upload file Template Word terlebih dahulu!');
      return;
    }
    setIsProcessing(true);
    try {
      const manualArray = manualNames.split('\n').map((n) => n.trim()).filter((n) => n);
      const allNames = [...new Set([...excelNames, ...manualArray])];
      
      if (allNames.length === 0) {
        alert('Harap masukkan setidaknya satu Nama Bank!');
        setIsProcessing(false);
        return;
      }

      const reader = new FileReader();
      reader.onload = async (event) => {
        try {
          const content = event.target.result;
          const zipResult = new JSZip();

          allNames.forEach((penerima) => {
            const zipTemplate = new PizZip(content);
            const doc = new Docxtemplater(zipTemplate, {
              paragraphLoop: true, linebreaks: true, delimiters: { start: '{{', end: '}}' }
            });
            
            // Map data sesuai placeholder dokumen
            const docData = { ...formData, nama_penerima: penerima };
            
            Object.keys(docData).forEach(key => {
               if(!docData[key]) docData[key] = `{{${key}}}`; 
               if((key === 'sebutan1' || key === 'sebutan2') && docData[key] === `{{${key}}}`) docData[key] = "";
            });
            
            doc.render(docData);
            const out = doc.getZip().generate({
              type: 'blob',
              mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            });
            
            if (allNames.length === 1) {
              setDownloadData({
                blob: out,
                fileName: `Konfirmasi Bank - ${penerima}.docx`,
                isZip: false
              });
            } else {
              zipResult.file(`Konfirmasi Bank - ${penerima}.docx`, out);
            }
          });
          
          if (allNames.length > 1) {
            const zipContent = await zipResult.generateAsync({ type: 'blob' });
            setDownloadData({
              blob: zipContent,
              fileName: `Konfirmasi Bank - ${formData.Nama_Klien || 'Klien'}.zip`,
              isZip: true
            });
          }
          setHasGenerated(true);
        } catch (error) {
          console.error(error);
          alert("Terjadi kesalahan saat memproses dokumen.");
        } finally {
          setIsProcessing(false);
        }
      };
      reader.readAsArrayBuffer(templateFile);
    } catch (error) {
      console.error(error);
      setIsProcessing(false);
    }
  };

  const totalRecipients = new Set([...excelNames, ...manualNames.split('\n').filter(n => n.trim())].filter(n => n)).size;
  const currentStep = hasGenerated ? 4 : activeStep;

  return (
    <div className="app-wrapper">
      <header className="app-header">
        <div className="app-header__logo">
          <span className="app-header__logo-icon"><Icons.Bank /></span>
          <h1 className="app-header__title">Generator Konfirmasi Bank</h1>
        </div>
        <p className="app-header__subtitle">
          Alat bantu audit untuk membuat surat konfirmasi bank secara massal dan otomatis.
        </p>
      </header>

      <main className="app-card">
        <div className="progress-bar">
          {[
            { step: 1, label: 'Template' },
            { step: 2, label: 'Detail Data' },
            { step: 3, label: 'Bank Penerima' },
          ].map(({ step, label }) => (
            <div key={step} className={`progress-bar__step ${step === currentStep ? 'active' : ''} ${step < currentStep ? 'completed' : ''}`}>
              <div className="progress-bar__circle">{step < currentStep ? <Icons.Check /> : step}</div>
              <span className="progress-bar__label">{label}</span>
            </div>
          ))}
        </div>

        {!hasGenerated ? (
          <>
            {/* Step 1: Template */}
            <section className="section">
              <div className="section__header">
                <span className="section__number">1</span>
                <h3 className="section__title">Upload Template</h3>
              </div>
              <p className="section__description">Pilih file template Surat Konfirmasi Bank (.docx)</p>
              
              <div className="file-upload mt-3">
                <label className="file-upload__area">
                  <input 
                    type="file" 
                    accept=".docx" 
                    className="file-upload__input"
                    onChange={(e) => {
                      setTemplateFile(e.target.files[0]);
                      setActiveStep(2);
                    }} 
                  />
                  <div className="file-upload__icon"><Icons.Upload /></div>
                  <p className="file-upload__text"><strong>Pilih file</strong> atau drag & drop</p>
                </label>
                {templateFile && (
                  <div className="file-upload__preview mt-2">
                    <Icons.File /> {templateFile.name}
                  </div>
                )}
              </div>

              {/* TOMBOL DOWNLOAD TEMPLATE BARU */}
              <div className="mt-3">
                <a 
                  href="/bahan/Konfirmasi-Bank-Template.docx" 
                  download 
                  className="btn btn--outline btn--full"
                >
                  <Icons.Download /> Download Template Standar Bank
                </a>
              </div>
            </section>

            {/* Step 2: Detail Dokumen */}
            <section className="section">
              <div className="section__header">
                <span className="section__number">2</span>
                <h3 className="section__title">Detail Dokumen</h3>
              </div>
              <div className="form-grid mt-3">
                <div className="form-row">
                  <div className="form-row__item">
                    <label className="form-label">Nama Klien <span className="form-label__required">*</span></label>
                    <input name="Nama_Klien" className="form-input" placeholder="PT Contoh" onChange={handleInputChange} value={formData.Nama_Klien} />
                  </div>
                  <div className="form-row__item">
                    <label className="form-label">Tanggal Tutup Buku <span className="form-label__required">*</span></label>
                    <input name="Tanggal_Konfirmasi" className="form-input" placeholder="31 Desember 2025" onChange={handleInputChange} value={formData.Tanggal_Konfirmasi} />
                  </div>
                </div>
                <div className="form-row">
                  <div className="form-row__item form-row__item--small">
                    <label className="form-label">Sebutan</label>
                    <input name="sebutan1" className="form-input" placeholder="Bpk" onChange={handleInputChange} value={formData.sebutan1} />
                  </div>
                  <div className="form-row__item">
                    <label className="form-label">Auditor 1</label>
                    <input name="Auditor1" className="form-input" placeholder="Nama Auditor" onChange={handleInputChange} value={formData.Auditor1} />
                  </div>
                </div>
                <div className="form-row">
                  <div className="form-row__item form-row__item--small">
                    <label className="form-label">Sebutan</label>
                    <input name="sebutan2" className="form-input" placeholder="Ibu" onChange={handleInputChange} value={formData.sebutan2} />
                  </div>
                  <div className="form-row__item">
                    <label className="form-label">Auditor 2</label>
                    <input name="Auditor2" className="form-input" placeholder="Nama Auditor" onChange={handleInputChange} value={formData.Auditor2} />
                  </div>
                </div>
                <div className="form-row">
                  <div className="form-row__item">
                    <label className="form-label">Penandatangan <span className="form-label__required">*</span></label>
                    <input name="Nama_Direktur" className="form-input" placeholder="Nama Direktur" onChange={handleInputChange} value={formData.Nama_Direktur} />
                  </div>
                  <div className="form-row__item">
                    <label className="form-label">Jabatan</label>
                    <input name="Jabatan" className="form-input" placeholder="Direktur Utama" onChange={handleInputChange} value={formData.Jabatan} />
                  </div>
                </div>
              </div>
            </section>

            {/* Step 3: Bank Penerima */}
            <section className="section">
              <div className="section__header">
                <span className="section__number">3</span>
                <h3 className="section__title">Bank Penerima</h3>
              </div>
              <div className="mt-3">
                <label className="form-label">Import Excel (Kolom "Nama Bank")</label>
                <div className="file-upload">
                  <label className="file-upload__area" style={{padding: '15px'}}>
                    <input type="file" accept=".xlsx, .xls" className="file-upload__input" onChange={handleExcelUpload} />
                    <p className="file-upload__text"><Icons.Users /> Klik untuk upload Excel</p>
                  </label>
                  {excelNames.length > 0 && <div className="file-upload__preview"><Icons.Check /> {excelNames.length} bank terdeteksi</div>}
                </div>
              </div>
              <div className="divider" />
              <label className="form-label">Input Manual (Satu bank per baris)</label>
              <textarea className="form-textarea" placeholder="Bank Mandiri KCP Samarinda..." value={manualNames} onChange={(e) => setManualNames(e.target.value)} />
              <p className="form-hint text-right">Total: <strong>{totalRecipients}</strong> bank</p>
            </section>

            <div className="action-bar">
              <button className="btn btn--ghost" onClick={() => window.location.reload()}>Reset</button>
              <button className="btn btn--primary btn--lg" onClick={generateDocuments} disabled={isProcessing || !templateFile}>
                {isProcessing ? 'Memproses...' : `Generate ${totalRecipients || 1} Dokumen`}
              </button>
            </div>
          </>
        ) : (
          <section className="section">
            <div className="result-card">
              <div className="result-card__icon"><Icons.Check /></div>
              <h4 className="result-card__title">Berhasil!</h4>
              <p className="result-card__message">{totalRecipients} Surat Konfirmasi Bank siap diunduh.</p>
              <button className="btn btn--success btn--lg" onClick={() => saveAs(downloadData.blob, downloadData.fileName)}>
                <Icons.Download /> Unduh {downloadData.isZip ? 'ZIP' : 'Dokumen'}
              </button>
            </div>
            <div className="action-bar mt-4">
              <button className="btn btn--ghost" onClick={() => setHasGenerated(false)}><Icons.ArrowLeft /> Kembali</button>
            </div>
          </section>
        )}
      </main>
    </div>
  );
}

export default App;