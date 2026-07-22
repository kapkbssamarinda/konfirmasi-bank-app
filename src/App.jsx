import { useState } from 'react';
import Swal from 'sweetalert2';
import Docxtemplater from 'docxtemplater';
import PizZip from 'pizzip';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';
import * as XLSX from 'xlsx';
import { motion, AnimatePresence } from 'motion/react';
import { useAutoAnimate } from '@formkit/auto-animate/react';
import ReactConfetti from 'react-confetti';
import { useWindowSize } from 'react-use';
import {
  Upload, FileText, Download, Check, ArrowLeft,
  Users, Building2, Loader2, FileDown,
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Textarea } from '@/components/ui/textarea';
import { Separator } from '@/components/ui/separator';
import logoTransparan from './assets/logo_transparan.png';

const STEPS = [
  { id: 1, label: 'Upload Template', icon: Upload },
  { id: 2, label: 'Detail Dokumen', icon: FileText },
  { id: 3, label: 'Bank Penerima', icon: Building2 },
];

function App() {
  const [templateFile, setTemplateFile] = useState(null);
  const [excelNames, setExcelNames] = useState([]);
  const [manualNames, setManualNames] = useState('');
  const [isProcessing, setIsProcessing] = useState(false);
  const [hasGenerated, setHasGenerated] = useState(false);
  const [downloadData, setDownloadData] = useState({ blob: null, fileName: '', isZip: false });
  const [activeStep, setActiveStep] = useState(1);
  const [showConfetti, setShowConfetti] = useState(false);
  const [outputFormat, setOutputFormat] = useState('docx');

  const [formData, setFormData] = useState({
    Tanggal_Konfirmasi: '',
    Nama_Klien: '',
    sebutan1: '',
    Auditor1: '',
    nomor_hp1: '',
    sebutan2: '',
    Auditor2: '',
    nomor_hp2: '',
    Nama_Direktur: '',
    Jabatan: ''
  });

  const [bankListRef] = useAutoAnimate();
  const { width, height } = useWindowSize();

  const handleInputChange = (e) => {
    const { name, value } = e.target;
    setFormData((prev) => ({ ...prev, [name]: value }));
  };

  const handlePhoneInputChange = (e) => {
    const { name } = e.target;
    let digits = e.target.value.replace(/\D/g, '');
    if (digits.startsWith('62')) digits = digits.slice(2);
    else if (digits.startsWith('0')) digits = digits.slice(1);
    digits = digits.slice(0, 11);
    const p1 = digits.slice(0, 3);
    const p2 = digits.slice(3, 7);
    const p3 = digits.slice(7, 11);
    let formatted = p1;
    if (p2) formatted += '-' + p2;
    if (p3) formatted += '-' + p3;
    setFormData((prev) => ({ ...prev, [name]: formatted }));
  };

  const convertDocxToPdf = async (docxBlob) => {
    // Konversi Blob → base64 untuk dikirim ke Vercel serverless function
    const base64 = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result.split(',')[1]);
      reader.onerror = reject;
      reader.readAsDataURL(docxBlob);
    });

    const response = await fetch('/api/convert', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ docxBase64: base64 }),
    });

    if (!response.ok) {
      const err = await response.json().catch(() => ({}));
      throw new Error(err.error || 'Gagal mengkonversi ke PDF');
    }

    return response.blob();
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
      Swal.fire({
        icon: 'success',
        title: 'Import Berhasil',
        text: `${names.length} nama bank berhasil diimport!`,
        confirmButtonColor: '#4f46e5',
        timer: 2500,
        timerProgressBar: true,
      });
    };
    reader.readAsArrayBuffer(file);
  };

  const generateDocuments = async () => {
    if (!templateFile) {
      await Swal.fire({
        icon: 'warning',
        title: 'Template Belum Diupload',
        text: 'Harap upload file Template Word terlebih dahulu!',
        confirmButtonColor: '#4f46e5',
        confirmButtonText: 'Mengerti',
      });
      return;
    }

    const requiredFields = [
      { key: 'Nama_Klien',         label: 'Nama Klien' },
      { key: 'Tanggal_Konfirmasi', label: 'Tanggal Tutup Buku' },
      { key: 'sebutan1',           label: 'Sebutan Auditor 1' },
      { key: 'Auditor1',           label: 'Auditor 1' },
      { key: 'nomor_hp1',          label: 'Nomor HP Auditor 1' },
      { key: 'sebutan2',           label: 'Sebutan Auditor 2' },
      { key: 'Auditor2',           label: 'Auditor 2' },
      { key: 'nomor_hp2',          label: 'Nomor HP Auditor 2' },
      { key: 'Nama_Direktur',      label: 'Penandatangan' },
      { key: 'Jabatan',            label: 'Jabatan' },
    ];
    const emptyFields = requiredFields.filter(f => !formData[f.key].trim());
    if (emptyFields.length > 0) {
      await Swal.fire({
        icon: 'warning',
        title: 'Kolom Belum Lengkap',
        html: `Harap lengkapi kolom berikut:<br><ul style="text-align:left;margin-top:8px;line-height:1.8">${emptyFields.map(f => `<li>${f.label}</li>`).join('')}</ul>`,
        confirmButtonColor: '#4f46e5',
        confirmButtonText: 'Mengerti',
      });
      return;
    }

    setIsProcessing(true);
    try {
      const manualArray = manualNames.split('\n').map((n) => n.trim()).filter((n) => n);
      const allNames = [...new Set([...excelNames, ...manualArray])];

      if (allNames.length === 0) {
        await Swal.fire({
          icon: 'warning',
          title: 'Bank Penerima Kosong',
          text: 'Harap masukkan setidaknya satu Nama Bank!',
          confirmButtonColor: '#4f46e5',
          confirmButtonText: 'Mengerti',
        });
        setIsProcessing(false);
        return;
      }

      const reader = new FileReader();
      reader.onload = async (event) => {
        try {
          const content = event.target.result;
          const zipResult = new JSZip();
          const ext = outputFormat === 'pdf' ? '.pdf' : '.docx';

          for (const penerima of allNames) {
            const zipTemplate = new PizZip(content);
            const doc = new Docxtemplater(zipTemplate, {
              paragraphLoop: true, linebreaks: true, delimiters: { start: '{{', end: '}}' }
            });

            const docData = { ...formData, nama_penerima: penerima };
            Object.keys(docData).forEach(key => {
              if (!docData[key]) docData[key] = `{{${key}}}`;
              if (['sebutan1', 'sebutan2', 'nomor_hp1', 'nomor_hp2'].includes(key) && docData[key] === `{{${key}}}`) docData[key] = '';
            });

            doc.render(docData);
            const out = doc.getZip().generate({
              type: 'blob',
              mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            });

            const finalBlob = outputFormat === 'pdf' ? await convertDocxToPdf(out) : out;

            if (allNames.length === 1) {
              setDownloadData({
                blob: finalBlob,
                fileName: `Konfirmasi Bank - ${penerima}${ext}`,
                isZip: false
              });
            } else {
              zipResult.file(`Konfirmasi Bank - ${penerima}${ext}`, finalBlob);
            }
          }

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
          Swal.fire({
            icon: 'error',
            title: 'Terjadi Kesalahan',
            text: 'Terjadi kesalahan saat memproses dokumen.',
            confirmButtonColor: '#4f46e5',
            confirmButtonText: 'Tutup',
          });
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

  const totalRecipients = new Set([
    ...excelNames,
    ...manualNames.split('\n').filter(n => n.trim())
  ].filter(n => n)).size;

  return (
    <div className="min-h-dvh" style={{ backgroundColor: '#f8fafc', fontFamily: "'Plus Jakarta Sans', system-ui, sans-serif" }}>
      {showConfetti && (
        <ReactConfetti
          width={width}
          height={height}
          recycle={false}
          numberOfPieces={300}
          colors={['#4f46e5', '#7c3aed', '#0ea5e9', '#10b981', '#f59e0b']}
        />
      )}

      {/* Header */}
      <header style={{ backgroundColor: '#0f172a', borderBottom: '1px solid #1e293b', position: 'sticky', top: 0, zIndex: 50 }}>
        <div style={{ maxWidth: '960px', margin: '0 auto', padding: '0 1.5rem', height: '56px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
            <img src={logoTransparan} alt="KAP Logo" style={{ height: '32px', width: '32px', objectFit: 'contain' }} />
            <div>
              <p style={{ fontSize: '10px', letterSpacing: '0.12em', color: '#64748b', textTransform: 'uppercase', fontWeight: 600, margin: 0, lineHeight: 1 }}>
                KAP Kuncara Budi Santosa &amp; Rekan
              </p>
              <h1 style={{ fontSize: '14px', fontWeight: 700, color: '#ffffff', margin: 0, lineHeight: 1.3 }}>
                Generator Konfirmasi Bank
              </h1>
            </div>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <span style={{ fontSize: '11px', color: '#475569', fontWeight: 500 }}>Alat Bantu Audit</span>
            <div style={{ width: '6px', height: '6px', borderRadius: '50%', backgroundColor: '#34d399' }} />
          </div>
        </div>
      </header>

      {/* Mobile step bar */}
      <div style={{ backgroundColor: '#ffffff', borderBottom: '1px solid #f1f5f9', position: 'sticky', top: '56px', zIndex: 40, display: 'block' }}>
        <div style={{ maxWidth: '960px', margin: '0 auto', padding: '10px 1.5rem' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
            {STEPS.map((step, i) => {
              const done = hasGenerated || activeStep > step.id;
              const active = !hasGenerated && activeStep === step.id;
              return (
                <div key={step.id} style={{ display: 'flex', alignItems: 'center', flex: i < STEPS.length - 1 ? undefined : 'none', gap: '6px' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                    <span style={{
                      width: '20px', height: '20px', borderRadius: '50%', display: 'flex', alignItems: 'center',
                      justifyContent: 'center', fontSize: '10px', fontWeight: 700, flexShrink: 0,
                      border: done ? 'none' : `1.5px solid ${active ? '#4f46e5' : '#cbd5e1'}`,
                      backgroundColor: done ? '#4f46e5' : active ? '#ffffff' : '#ffffff',
                      color: done ? '#ffffff' : active ? '#4f46e5' : '#94a3b8',
                    }}>
                      {done ? <Check size={10} /> : step.id}
                    </span>
                    <span style={{ fontSize: '11px', fontWeight: 600, color: active ? '#4338ca' : done ? '#475569' : '#94a3b8', whiteSpace: 'nowrap' }}>
                      {step.label}
                    </span>
                  </div>
                  {i < STEPS.length - 1 && (
                    <div style={{ flex: 1, minWidth: '16px', height: '1px', backgroundColor: done ? '#a5b4fc' : '#e2e8f0', marginLeft: '6px' }} />
                  )}
                </div>
              );
            })}
          </div>
        </div>
      </div>

      {/* Main */}
      <main style={{ maxWidth: '960px', margin: '0 auto', padding: '2rem 1.5rem' }}>
        <div style={{ display: 'grid', gridTemplateColumns: '180px 1fr', gap: '2.5rem', alignItems: 'start' }}>

          {/* Desktop sidebar stepper */}
          <aside style={{ position: 'sticky', top: 'calc(56px + 44px + 1.5rem)' }}>
            <p style={{ fontSize: '10px', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.15em', color: '#94a3b8', marginBottom: '16px', paddingLeft: '12px' }}>
              Langkah
            </p>
            <nav>
              {STEPS.map((step, i) => {
                const done = hasGenerated || activeStep > step.id;
                const active = !hasGenerated && activeStep === step.id;
                return (
                  <div key={step.id} style={{ position: 'relative' }}>
                    {i < STEPS.length - 1 && (
                      <div style={{
                        position: 'absolute', left: '21px', top: '36px', width: '1px', height: '24px',
                        backgroundColor: done ? '#a5b4fc' : '#e2e8f0'
                      }} />
                    )}
                    <div style={{
                      display: 'flex', alignItems: 'center', gap: '12px', padding: '8px 12px', borderRadius: '8px',
                      backgroundColor: active ? '#eef2ff' : 'transparent', transition: 'background 0.15s',
                      marginBottom: '4px'
                    }}>
                      <span style={{
                        width: '24px', height: '24px', borderRadius: '50%', display: 'flex', alignItems: 'center',
                        justifyContent: 'center', fontSize: '11px', fontWeight: 700, flexShrink: 0,
                        outline: `1.5px solid ${done ? '#4f46e5' : active ? '#4f46e5' : '#cbd5e1'}`,
                        outlineOffset: '0px',
                        backgroundColor: done ? '#4f46e5' : '#ffffff',
                        color: done ? '#ffffff' : active ? '#4f46e5' : '#94a3b8',
                      }}>
                        {done ? <Check size={12} /> : step.id}
                      </span>
                      <span style={{ fontSize: '12px', fontWeight: 600, color: active ? '#4338ca' : done ? '#475569' : '#94a3b8' }}>
                        {step.label}
                      </span>
                    </div>
                  </div>
                );
              })}
            </nav>
          </aside>

          {/* Content */}
          <div>
            <AnimatePresence mode="wait">
              {!hasGenerated ? (
                <motion.div key="form"
                  initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0, y: -16 }}
                  transition={{ duration: 0.25 }}
                  style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}
                >
                  {/* Card 1: Upload Template */}
                  <Card>
                    <CardHeader>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                        <div style={{ width: '28px', height: '28px', borderRadius: '8px', backgroundColor: '#eef2ff', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                          <Upload size={15} color="#4f46e5" />
                        </div>
                        <div>
                          <CardTitle style={{ fontSize: '14px' }}>Upload Template</CardTitle>
                          <p style={{ fontSize: '12px', color: '#64748b', margin: '2px 0 0' }}>Pilih file template Surat Konfirmasi Bank (.docx)</p>
                        </div>
                      </div>
                    </CardHeader>
                    <CardContent style={{ display: 'flex', flexDirection: 'column', gap: '12px' }}>
                      <label style={{
                        display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center',
                        border: '2px dashed #e2e8f0', borderRadius: '12px', padding: '32px 24px', cursor: 'pointer',
                        transition: 'all 0.2s',
                      }}
                        onMouseEnter={e => { e.currentTarget.style.borderColor = '#a5b4fc'; e.currentTarget.style.backgroundColor = '#fafbff'; }}
                        onMouseLeave={e => { e.currentTarget.style.borderColor = '#e2e8f0'; e.currentTarget.style.backgroundColor = 'transparent'; }}
                      >
                        <input type="file" accept=".docx" style={{ display: 'none' }}
                          onChange={(e) => { setTemplateFile(e.target.files[0]); setActiveStep(2); }} />
                        <Upload size={28} color="#cbd5e1" style={{ marginBottom: '8px' }} />
                        <p style={{ fontSize: '13px', fontWeight: 600, color: '#475569', margin: '0 0 4px' }}>Pilih file atau seret ke sini</p>
                        <p style={{ fontSize: '11px', color: '#94a3b8', margin: 0 }}>.docx · Maks. 10 MB</p>
                      </label>

                      {templateFile && (
                        <motion.div initial={{ opacity: 0, y: 4 }} animate={{ opacity: 1, y: 0 }}
                          style={{ display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: '#eef2ff', color: '#4338ca', borderRadius: '8px', padding: '8px 12px', fontSize: '13px', fontWeight: 500 }}>
                          <FileText size={14} style={{ flexShrink: 0 }} />
                          <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{templateFile.name}</span>
                        </motion.div>
                      )}

                      <Button variant="outline" size="sm" asChild>
                        <a href="/bahan/Konfirmasi-Bank-Template.docx" download style={{ textDecoration: 'none' }}>
                          <Download size={13} style={{ marginRight: '6px' }} />
                          Download Template Standar
                        </a>
                      </Button>
                    </CardContent>
                  </Card>

                  {/* Card 2: Detail Dokumen */}
                  <Card>
                    <CardHeader>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                        <div style={{ width: '28px', height: '28px', borderRadius: '8px', backgroundColor: '#f5f3ff', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                          <FileText size={15} color="#7c3aed" />
                        </div>
                        <div>
                          <CardTitle style={{ fontSize: '14px' }}>Detail Dokumen</CardTitle>
                          <p style={{ fontSize: '12px', color: '#64748b', margin: '2px 0 0' }}>Lengkapi informasi yang akan terisi di setiap surat</p>
                        </div>
                      </div>
                    </CardHeader>
                    <CardContent>
                      {/* Klien + Tanggal */}
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px', marginBottom: '16px' }}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                          <Label>Nama Klien <span style={{ color: '#f43f5e' }}>*</span></Label>
                          <Input name="Nama_Klien" placeholder="PT Contoh" onChange={handleInputChange} value={formData.Nama_Klien} />
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                          <Label>Tanggal Tutup Buku <span style={{ color: '#f43f5e' }}>*</span></Label>
                          <Input name="Tanggal_Konfirmasi" placeholder="31 Desember 2025" onChange={handleInputChange} value={formData.Tanggal_Konfirmasi} />
                        </div>
                      </div>

                      <Separator style={{ marginBottom: '16px' }} />

                      {/* Auditor 1 */}
                      <p style={{ fontSize: '10px', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.12em', color: '#94a3b8', margin: '0 0 10px' }}>Auditor 1</p>
                      <div style={{ display: 'grid', gridTemplateColumns: '80px 1fr 1fr', gap: '10px', marginBottom: '16px' }}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                          <Label>Sebutan</Label>
                          <Input name="sebutan1" placeholder="Bpk" onChange={handleInputChange} value={formData.sebutan1} />
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                          <Label>Nama <span style={{ color: '#f43f5e' }}>*</span></Label>
                          <Input name="Auditor1" placeholder="Nama Auditor" onChange={handleInputChange} value={formData.Auditor1} />
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                          <Label>Nomor HP <span style={{ color: '#f43f5e' }}>*</span></Label>
                          <Input name="nomor_hp1" placeholder="822-5291-6183" inputMode="numeric" maxLength={13} onChange={handlePhoneInputChange} value={formData.nomor_hp1} />
                        </div>
                      </div>

                      {/* Auditor 2 */}
                      <p style={{ fontSize: '10px', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.12em', color: '#94a3b8', margin: '0 0 10px' }}>Auditor 2</p>
                      <div style={{ display: 'grid', gridTemplateColumns: '80px 1fr 1fr', gap: '10px', marginBottom: '16px' }}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                          <Label>Sebutan</Label>
                          <Input name="sebutan2" placeholder="Ibu" onChange={handleInputChange} value={formData.sebutan2} />
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                          <Label>Nama <span style={{ color: '#f43f5e' }}>*</span></Label>
                          <Input name="Auditor2" placeholder="Nama Auditor" onChange={handleInputChange} value={formData.Auditor2} />
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                          <Label>Nomor HP <span style={{ color: '#f43f5e' }}>*</span></Label>
                          <Input name="nomor_hp2" placeholder="812-3456-7890" inputMode="numeric" maxLength={13} onChange={handlePhoneInputChange} value={formData.nomor_hp2} />
                        </div>
                      </div>

                      <Separator style={{ marginBottom: '16px' }} />

                      {/* Penandatangan + Jabatan */}
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                          <Label>Penandatangan <span style={{ color: '#f43f5e' }}>*</span></Label>
                          <Input name="Nama_Direktur" placeholder="Nama Direktur" onChange={handleInputChange} value={formData.Nama_Direktur} />
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                          <Label>Jabatan <span style={{ color: '#f43f5e' }}>*</span></Label>
                          <Input name="Jabatan" placeholder="Direktur Utama" onChange={handleInputChange} value={formData.Jabatan} />
                        </div>
                      </div>
                    </CardContent>
                  </Card>

                  {/* Card 3: Bank Penerima */}
                  <Card>
                    <CardHeader>
                      <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                        <div style={{ width: '28px', height: '28px', borderRadius: '8px', backgroundColor: '#ecfdf5', display: 'flex', alignItems: 'center', justifyContent: 'center', flexShrink: 0 }}>
                          <Building2 size={15} color="#059669" />
                        </div>
                        <div>
                          <CardTitle style={{ fontSize: '14px' }}>Bank Penerima</CardTitle>
                          <p style={{ fontSize: '12px', color: '#64748b', margin: '2px 0 0' }}>Import dari Excel atau input nama bank secara manual</p>
                        </div>
                      </div>
                    </CardHeader>
                    <CardContent style={{ display: 'flex', flexDirection: 'column', gap: '14px' }}>
                      {/* Excel upload */}
                      <div>
                        <Label style={{ display: 'block', marginBottom: '6px' }}>Import Excel (Kolom "Nama Bank")</Label>
                        <label style={{
                          display: 'flex', alignItems: 'center', gap: '12px', border: '1px solid #e2e8f0',
                          borderRadius: '8px', padding: '10px 14px', cursor: 'pointer', transition: 'all 0.15s'
                        }}
                          onMouseEnter={e => { e.currentTarget.style.backgroundColor = '#f8fafc'; e.currentTarget.style.borderColor = '#6ee7b7'; }}
                          onMouseLeave={e => { e.currentTarget.style.backgroundColor = 'transparent'; e.currentTarget.style.borderColor = '#e2e8f0'; }}
                        >
                          <input type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={handleExcelUpload} />
                          <Users size={15} color="#94a3b8" style={{ flexShrink: 0 }} />
                          <span style={{ fontSize: '13px', color: '#64748b' }}>Klik untuk upload file Excel</span>
                        </label>
                        <div ref={bankListRef}>
                          {excelNames.length > 0 && (
                            <div style={{ marginTop: '8px', display: 'flex', alignItems: 'center', gap: '8px', backgroundColor: '#ecfdf5', color: '#047857', borderRadius: '8px', padding: '8px 12px', fontSize: '13px', fontWeight: 500 }}>
                              <Check size={14} style={{ flexShrink: 0 }} />
                              <span>{excelNames.length} bank terdeteksi dari Excel</span>
                            </div>
                          )}
                        </div>
                      </div>

                      {/* Divider */}
                      <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                        <Separator style={{ flex: 1 }} />
                        <span style={{ fontSize: '10px', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.1em', color: '#94a3b8', whiteSpace: 'nowrap' }}>ATAU</span>
                        <Separator style={{ flex: 1 }} />
                      </div>

                      {/* Manual input */}
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                        <Label>Input Manual (Satu bank per baris)</Label>
                        <Textarea
                          placeholder={'Bank Mandiri KCP Samarinda\nBank BRI KCP Balikpapan\n...'}
                          style={{ minHeight: '88px', fontFamily: 'monospace', fontSize: '13px' }}
                          value={manualNames}
                          onChange={(e) => setManualNames(e.target.value)}
                        />
                      </div>

                      {/* Total badge */}
                      <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
                        <span style={{
                          display: 'inline-flex', alignItems: 'center', gap: '6px', padding: '4px 10px',
                          borderRadius: '9999px', fontSize: '12px', fontWeight: 600,
                          backgroundColor: totalRecipients > 0 ? '#ecfdf5' : '#f1f5f9',
                          color: totalRecipients > 0 ? '#047857' : '#94a3b8',
                          border: `1px solid ${totalRecipients > 0 ? '#a7f3d0' : '#e2e8f0'}`,
                        }}>
                          {totalRecipients > 0 ? <Check size={11} /> : <span>○</span>}
                          Total: {totalRecipients} bank
                        </span>
                      </div>
                    </CardContent>
                  </Card>

                  {/* Action bar */}
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', paddingTop: '8px', gap: '12px' }}>
                    <Button variant="ghost" size="sm" onClick={() => window.location.reload()}>
                      Reset
                    </Button>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                      {/* Format toggle */}
                      <div style={{ display: 'flex', borderRadius: '6px', overflow: 'hidden', border: '1px solid #e2e8f0' }}>
                        {['docx', 'pdf'].map(fmt => (
                          <button key={fmt}
                            style={{
                              padding: '5px 12px', fontSize: '11px', fontWeight: 700, letterSpacing: '0.05em',
                              border: 'none', cursor: 'pointer', transition: 'all 0.15s',
                              backgroundColor: outputFormat === fmt ? '#0f172a' : '#ffffff',
                              color: outputFormat === fmt ? '#ffffff' : '#64748b',
                              fontFamily: 'inherit',
                            }}
                            onClick={() => setOutputFormat(fmt)}
                          >{fmt.toUpperCase()}</button>
                        ))}
                      </div>

                      <button
                        onClick={generateDocuments}
                        disabled={isProcessing || !templateFile}
                        style={{
                          display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
                          gap: '8px', padding: '0 20px', height: '40px', borderRadius: '8px',
                          backgroundColor: isProcessing || !templateFile ? '#a5b4fc' : '#4f46e5',
                          color: '#ffffff', fontSize: '14px', fontWeight: 600, border: 'none',
                          cursor: isProcessing || !templateFile ? 'not-allowed' : 'pointer',
                          boxShadow: '0 1px 3px rgba(79,70,229,0.3)',
                          transition: 'all 0.15s', fontFamily: 'inherit', minWidth: '160px',
                        }}
                        onMouseEnter={e => { if (!isProcessing && templateFile) e.currentTarget.style.backgroundColor = '#4338ca'; }}
                        onMouseLeave={e => { if (!isProcessing && templateFile) e.currentTarget.style.backgroundColor = '#4f46e5'; }}
                      >
                        {isProcessing ? (
                          <><Loader2 size={15} style={{ animation: 'spin 1s linear infinite' }} />Memproses...</>
                        ) : (
                          `Generate ${totalRecipients || 1} Dokumen`
                        )}
                      </button>
                    </div>
                  </div>
                </motion.div>
              ) : (
                /* Success state */
                <motion.div key="result"
                  initial={{ opacity: 0, scale: 0.96 }} animate={{ opacity: 1, scale: 1 }}
                  transition={{ duration: 0.35, ease: [0.34, 1.56, 0.64, 1] }}
                >
                  <Card>
                    <CardContent style={{ padding: '64px 32px', display: 'flex', flexDirection: 'column', alignItems: 'center', textAlign: 'center', gap: '20px' }}>
                      <motion.div
                        initial={{ scale: 0 }} animate={{ scale: 1 }}
                        transition={{ delay: 0.15, duration: 0.4, ease: [0.34, 1.56, 0.64, 1] }}
                        style={{ width: '64px', height: '64px', borderRadius: '50%', backgroundColor: '#dcfce7', display: 'flex', alignItems: 'center', justifyContent: 'center' }}
                      >
                        <Check size={30} color="#16a34a" />
                      </motion.div>
                      <div>
                        <h2 style={{ fontSize: '24px', fontWeight: 700, color: '#0f172a', margin: '0 0 6px' }}>Berhasil!</h2>
                        <p style={{ fontSize: '14px', color: '#64748b', margin: '0 0 4px' }}>{totalRecipients} Surat Konfirmasi Bank siap diunduh.</p>
                        <p style={{ fontSize: '12px', color: '#94a3b8', fontFamily: 'monospace', margin: 0 }}>{downloadData.fileName}</p>
                      </div>
                      <button
                        onClick={() => saveAs(downloadData.blob, downloadData.fileName)}
                        style={{
                          display: 'inline-flex', alignItems: 'center', gap: '8px', padding: '0 24px', height: '44px',
                          borderRadius: '8px', backgroundColor: '#16a34a', color: '#ffffff', fontSize: '14px',
                          fontWeight: 600, border: 'none', cursor: 'pointer', boxShadow: '0 1px 3px rgba(22,163,74,0.3)',
                          fontFamily: 'inherit',
                        }}
                      >
                        <FileDown size={16} />
                        Unduh {downloadData.isZip ? 'ZIP' : outputFormat.toUpperCase()}
                      </button>
                    </CardContent>
                  </Card>

                  <div style={{ marginTop: '16px' }}>
                    <Button variant="ghost" size="sm"
                      onClick={() => { setHasGenerated(false); setShowConfetti(false); }}>
                      <ArrowLeft size={14} style={{ marginRight: '6px' }} /> Buat lagi
                    </Button>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </div>
        </div>
      </main>

      <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
    </div>
  );
}

export default App;
