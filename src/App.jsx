import { useState } from 'react';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';
import { motion, AnimatePresence } from 'motion/react';
import { useAutoAnimate } from '@formkit/auto-animate/react';
import ReactConfetti from 'react-confetti';
import { useWindowSize } from 'react-use';
import logoTransparan from './assets/logo_transparan.png';
import logo from './assets/logo.png';
import './App.css';

const Icons = {
  Bank: () => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" width="24" height="24">
      <path strokeLinecap="round" strokeLinejoin="round" d="M12 21v-8.25M15.75 21v-8.25M8.25 21v-8.25M3 9l9-6 9 6m-1.5 12V10.332A48.36 48.36 0 0012 9.75c-2.551 0-5.056.2-7.5.582V21M3 21h18M12 6.75h.008v.008H12V6.75z" />
    </svg>
  ),
  Users: () => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" width="18" height="18">
      <path strokeLinecap="round" strokeLinejoin="round" d="M15 19.128a9.38 9.38 0 002.625.372 9.337 9.337 0 004.121-.952 4.125 4.125 0 00-7.533-2.493M15 19.128v-.003c0-1.113-.285-2.16-.786-3.07M15 19.128v.106A12.318 12.318 0 018.624 21c-2.331 0-4.512-.645-6.374-1.766l-.001-.109a6.375 6.375 0 0111.964-3.07M12 6.375a3.375 3.375 0 11-6.75 0 3.375 3.375 0 016.75 0zm8.25 2.25a2.625 2.625 0 11-5.25 0 2.625 2.625 0 015.25 0z" />
    </svg>
  ),
  Download: () => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" width="18" height="18">
      <path strokeLinecap="round" strokeLinejoin="round" d="M3 16.5v2.25A2.25 2.25 0 005.25 21h13.5A2.25 2.25 0 0021 18.75V16.5M16.5 12L12 16.5m0 0L7.5 12m4.5 4.5V3" />
    </svg>
  ),
  Check: () => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={2.5} stroke="currentColor" width="16" height="16">
      <path strokeLinecap="round" strokeLinejoin="round" d="M4.5 12.75l6 6 9-13.5" />
    </svg>
  ),
  ArrowLeft: () => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" width="18" height="18">
      <path strokeLinecap="round" strokeLinejoin="round" d="M10.5 19.5L3 12m0 0l7.5-7.5M3 12h18" />
    </svg>
  ),
  File: () => (
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" strokeWidth={1.5} stroke="currentColor" width="18" height="18">
      <path strokeLinecap="round" strokeLinejoin="round" d="M19.5 14.25v-2.625a3.375 3.375 0 00-3.375-3.375h-1.5A1.125 1.125 0 0113.5 7.125v-1.5a3.375 3.375 0 00-3.375-3.375H8.25m2.25 0H5.625c-.621 0-1.125.504-1.125 1.125v17.25c0 .621.504 1.125 1.125 1.125h12.75c.621 0 1.125-.504 1.125-1.125V11.25a9 9 0 00-9-9z" />
    </svg>
  ),
};

function App() {
  const [templateFile, setTemplateFile] = useState(null);
  const [excelNames, setExcelNames] = useState([]);
  const [manualNames, setManualNames] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [hasGenerated, setHasGenerated] = useState(false);
  const [downloadData, setDownloadData] = useState({ blob: null, fileName: '', isZip: false });
  const [activeStep, setActiveStep] = useState(1);
  const [showConfetti, setShowConfetti] = useState(false);

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

  const [bankListRef] = useAutoAnimate();
  const { width, height } = useWindowSize();

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
              if (!docData[key]) docData[key] = `{{${key}}}`;
              if ((key === 'sebutan1' || key === 'sebutan2') && docData[key] === `{{${key}}}`) docData[key] = "";
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
          setShowConfetti(true);
          setTimeout(() => setShowConfetti(false), 5000);
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
        <motion.div
          initial={{ opacity: 0, y: -20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.6, ease: 'easeOut' }}
        >
          <img src={logoTransparan} alt="KAP Kuncara Budi Santosa & Rekan" className="app-header__logo-img" />
          <div className="app-header__kap-badge">
            KAP Kuncara Budi Santosa &amp; Rekan
          </div>
          <h1 className="app-header__title">Generator Konfirmasi Bank</h1>
          <p className="app-header__subtitle">
            Alat bantu audit untuk membuat surat konfirmasi bank secara massal dan otomatis.
          </p>
        </motion.div>
      </header>

      <motion.main
        className="app-card"
        initial={{ opacity: 0, y: 30 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.5, delay: 0.2, ease: 'easeOut' }}
      >
        {/* Progress Bar */}
        <div className="progress-bar">
          {[
            { step: 1, label: 'Template' },
            { step: 2, label: 'Detail Data' },
            { step: 3, label: 'Bank Penerima' },
          ].map(({ step, label }) => (
            <div
              key={step}
              className={`progress-bar__step ${step === currentStep ? 'active' : ''} ${step < currentStep ? 'completed' : ''}`}
            >
              <div className="progress-bar__circle">
                {step < currentStep ? <Icons.Check /> : step}
              </div>
              <span className="progress-bar__label">{label}</span>
            </div>
          ))}
        </div>

        <AnimatePresence mode="wait">
          {!hasGenerated ? (
            <motion.div
              key="form"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0, y: -20 }}
              transition={{ duration: 0.3 }}
            >
              {/* Step 1: Template */}
              <motion.section
                className="section"
                initial={{ opacity: 0, x: -20 }}
                animate={{ opacity: 1, x: 0 }}
                transition={{ duration: 0.4, delay: 0.1 }}
              >
                <div className="section__header">
                  <span className="section__number">1</span>
                  <div>
                    <h3 className="section__title">Upload Template</h3>
                    <p className="section__description">Pilih file template Surat Konfirmasi Bank (.docx)</p>
                  </div>
                </div>

                <div className="file-upload mt-3">
                  <motion.label
                    className="file-upload__area"
                    whileHover={{ scale: 1.01 }}
                    whileTap={{ scale: 0.99 }}
                  >
                    <input
                      type="file"
                      accept=".docx"
                      className="file-upload__input"
                      onChange={(e) => {
                        setTemplateFile(e.target.files[0]);
                        setActiveStep(2);
                      }}
                    />
                    <div className="file-upload__icon">📤</div>
                    <p className="file-upload__text"><strong>Pilih file</strong> atau drag &amp; drop</p>
                    <p className="file-upload__hint">.docx &bull; Maks. 10MB</p>
                  </motion.label>
                  {templateFile && (
                    <div className="file-upload__preview">
                      <Icons.File /> {templateFile.name}
                    </div>
                  )}
                </div>

                <div className="mt-3">
                  <a
                    href="/bahan/Konfirmasi-Bank-Template.docx"
                    download
                    className="btn btn--outline"
                  >
                    <Icons.Download /> Download Template Standar Bank
                  </a>
                </div>
              </motion.section>

              {/* Step 2: Detail Dokumen */}
              <motion.section
                className="section"
                initial={{ opacity: 0, x: -20 }}
                animate={{ opacity: 1, x: 0 }}
                transition={{ duration: 0.4, delay: 0.2 }}
              >
                <div className="section__header">
                  <span className="section__number">2</span>
                  <div>
                    <h3 className="section__title">Detail Dokumen</h3>
                    <p className="section__description">Lengkapi informasi yang akan terisi di setiap surat</p>
                  </div>
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
                  <div className="auditor-divider" />
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
              </motion.section>

              {/* Step 3: Bank Penerima */}
              <motion.section
                className="section"
                initial={{ opacity: 0, x: -20 }}
                animate={{ opacity: 1, x: 0 }}
                transition={{ duration: 0.4, delay: 0.3 }}
              >
                <div className="section__header">
                  <span className="section__number">3</span>
                  <div>
                    <h3 className="section__title">Bank Penerima</h3>
                    <p className="section__description">Import dari Excel atau input nama bank secara manual</p>
                  </div>
                </div>
                <div className="mt-3">
                  <label className="form-label">Import Excel (Kolom "Nama Bank")</label>
                  <div className="file-upload">
                    <motion.label
                      className="file-upload__area file-upload__area--compact"
                      whileHover={{ scale: 1.01 }}
                      whileTap={{ scale: 0.99 }}
                    >
                      <input type="file" accept=".xlsx, .xls" className="file-upload__input" onChange={handleExcelUpload} />
                      <p className="file-upload__text">
                        <Icons.Users />&nbsp;&nbsp;Klik untuk upload Excel
                      </p>
                    </motion.label>
                    <div ref={bankListRef}>
                      {excelNames.length > 0 && (
                        <div className="file-upload__preview">
                          <Icons.Check /> {excelNames.length} bank terdeteksi dari Excel
                        </div>
                      )}
                    </div>
                  </div>
                </div>
                <div className="or-separator">ATAU</div>
                <label className="form-label">Input Manual (Satu bank per baris)</label>
                <textarea
                  className="form-textarea"
                  placeholder="Bank Mandiri KCP Samarinda..."
                  value={manualNames}
                  onChange={(e) => setManualNames(e.target.value)}
                />
                <div style={{ textAlign: 'right', marginTop: '8px' }}>
                  <span className={`bank-count-badge ${totalRecipients > 0 ? 'bank-count-badge--active' : 'bank-count-badge--zero'}`}>
                    {totalRecipients > 0 ? '✓' : '○'} Total: {totalRecipients} bank
                  </span>
                </div>
              </motion.section>

              <div className="action-bar">
                <button className="btn btn--ghost" onClick={() => window.location.reload()}>Reset</button>
                <div className="action-bar__end">
                  {!templateFile && (
                    <span className="action-bar__hint">Upload template terlebih dahulu</span>
                  )}
                  <motion.button
                    className="btn btn--primary btn--lg"
                    onClick={generateDocuments}
                    disabled={isProcessing || !templateFile}
                    whileHover={{ scale: 1.03 }}
                    whileTap={{ scale: 0.97 }}
                  >
                    {isProcessing ? 'Memproses...' : `Generate ${totalRecipients || 1} Dokumen`}
                  </motion.button>
                </div>
              </div>
            </motion.div>
          ) : (
            <motion.div
              key="result"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              transition={{ duration: 0.3 }}
            >
              {showConfetti && (
                <ReactConfetti
                  width={width}
                  height={height}
                  recycle={false}
                  numberOfPieces={300}
                  colors={['#667eea', '#764ba2', '#f093fb', '#4facfe', '#00c9a7']}
                />
              )}
              <motion.section
                className="section"
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                transition={{ duration: 0.5, ease: [0.34, 1.56, 0.64, 1] }}
              >
                <div className="result-card">
                  <motion.div
                    className="result-card__icon"
                    initial={{ scale: 0 }}
                    animate={{ scale: 1 }}
                    transition={{ delay: 0.2, duration: 0.5, ease: [0.34, 1.56, 0.64, 1] }}
                  >
                    <Icons.Check />
                  </motion.div>
                  <h4 className="result-card__title">Berhasil! 🎉</h4>
                  <p className="result-card__message">{totalRecipients} Surat Konfirmasi Bank siap diunduh.</p>
                  <p className="result-card__filename">{downloadData.fileName}</p>
                  <motion.button
                    className="btn btn--success btn--lg"
                    onClick={() => saveAs(downloadData.blob, downloadData.fileName)}
                    whileHover={{ scale: 1.05 }}
                    whileTap={{ scale: 0.95 }}
                  >
                    <Icons.Download /> Unduh {downloadData.isZip ? 'ZIP' : 'Dokumen'}
                  </motion.button>
                </div>
                <div className="action-bar mt-4">
                  <button
                    className="btn btn--ghost"
                    onClick={() => { setHasGenerated(false); setShowConfetti(false); }}
                  >
                    <Icons.ArrowLeft /> Buat lagi
                  </button>
                </div>
              </motion.section>
            </motion.div>
          )}
        </AnimatePresence>
      </motion.main>
    </div>
  );
}

export default App;
