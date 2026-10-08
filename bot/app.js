/**
 * Word to PDF Converter - Core Application Logic
 * Supports docx-preview, mammoth fallback, html2pdf rendering, and native Vector Print-to-PDF
 */

(function () {
  'use strict';

  // --- Multi-language Dictionary ---
  const i18n = {
    uz: {
      heroPill: "⚡ 100% Xavfsiz & Mahalliy (Brauzer ichida)",
      heroTitle: "Word hujjatlarini <span class='gradient-text'>PDF</span> formatiga aylantiring",
      heroSubtitle: "Hech qanday serverga yuklanmaydi, hujjatlaringiz xavfsiz va faqat kompyuteringizda soniyalar ichida yuqori sifatli PDF ga aylanadi.",
      dropTitle: "Word faylini shu yerga tashlang",
      dropDesc: "yoki kompyuterdan tanlash uchun bosing (.docx, .doc)",
      browseBtn: "Faylni tanlash",
      optSize: "Sahifa o'lchami",
      optOrientation: "Sahifa yo'nalishi",
      optPortrait: "Vertikal (Tik)",
      optLandscape: "Gorizontal (Yotiq)",
      optMargins: "Hoshiyalar (Margin)",
      optMarginStd: "Standart (10mm)",
      optMarginNone: "Hoshiyasiz (0mm)",
      optMarginNarrow: "Tor (5mm)",
      optMarginWide: "Keng (20mm)",
      optQuality: "Sifat (DPI)",
      optQualityStd: "Standart (Yuqori tezlik)",
      optQualityHigh: "Maksimal (300 DPI)",
      convertBtn: "PDF qilib yuklab olish",
      printBtn: "Chop etish / Vektor PDF",
      converting: "PDF tayyorlanmoqda...",
      ready: "PDF tayyor!",
      readingFile: "Word fayli tahlil qilinmoqda...",
      renderingPages: "Sahifalar tuzilmoqda...",
      generatingPdf: "PDF fayli yaratilmoqda...",
      successToast: "Hujjat PDF ga muvaffaqiyatli aylantirildi!",
      errorToast: "Xatolik yuz berdi. Iltimos faylni tekshiring.",
      invalidFileType: "Iltimos faqat .docx yoki .doc formatidagi Word fayllarini yuklang!",
      previewTitle: "Hujjatni oldindan ko'rish",
      zoomIn: "Kattalashtirish",
      zoomOut: "Kichiklashtirish",
      removeBtn: "Boshqa fayl tanlash"
    },
    ru: {
      heroPill: "⚡ 100% Безопасно & Локально в браузере",
      heroTitle: "Конвертируйте Word в <span class='gradient-text'>PDF</span>",
      heroSubtitle: "Файлы не загружаются на сервер. Конвертация происходит прямо на вашем компьютере за секунды.",
      dropTitle: "Перетащите Word файл сюда",
      dropDesc: "или нажмите для выбора с компьютера (.docx, .doc)",
      browseBtn: "Выбрать файл",
      optSize: "Размер страницы",
      optOrientation: "Ориентация",
      optPortrait: "Книжная (Вертикальная)",
      optLandscape: "Альбомная (Горизонтальная)",
      optMargins: "Поля",
      optMarginStd: "Стандартные (10мм)",
      optMarginNone: "Без полей (0мм)",
      optMarginNarrow: "Узкие (5мм)",
      optMarginWide: "Широкие (20мм)",
      optQuality: "Качество (DPI)",
      optQualityStd: "Стандартное (Быстро)",
      optQualityHigh: "Максимальное (300 DPI)",
      convertBtn: "Скачать как PDF",
      printBtn: "Печать / Векторный PDF",
      converting: "Создание PDF...",
      ready: "PDF готов!",
      readingFile: "Чтение файла...",
      renderingPages: "Построение страниц...",
      generatingPdf: "Генерация PDF...",
      successToast: "Документ успешно конвертирован в PDF!",
      errorToast: "Произошла ошибка при обработке файла.",
      invalidFileType: "Пожалуйста, выберите файл формата .docx или .doc!",
      previewTitle: "Предпросмотр документа",
      zoomIn: "Увеличить",
      zoomOut: "Уменьшить",
      removeBtn: "Выбрать другой файл"
    },
    en: {
      heroPill: "⚡ 100% Private & Client-side Conversion",
      heroTitle: "Convert Word to <span class='gradient-text'>PDF</span> Instantly",
      heroSubtitle: "No files are uploaded to any server. Your documents are safely processed right inside your browser in seconds.",
      dropTitle: "Drag & drop Word document here",
      dropDesc: "or click to browse from computer (.docx, .doc)",
      browseBtn: "Choose File",
      optSize: "Page Size",
      optOrientation: "Orientation",
      optPortrait: "Portrait",
      optLandscape: "Landscape",
      optMargins: "Margins",
      optMarginStd: "Standard (10mm)",
      optMarginNone: "No Margins (0mm)",
      optMarginNarrow: "Narrow (5mm)",
      optMarginWide: "Wide (20mm)",
      optQuality: "Quality (DPI)",
      optQualityStd: "Standard (Fast)",
      optQualityHigh: "Ultra High (300 DPI)",
      convertBtn: "Download PDF",
      printBtn: "Print / Vector PDF",
      converting: "Generating PDF...",
      ready: "PDF Ready!",
      readingFile: "Parsing Word document...",
      renderingPages: "Rendering layout...",
      generatingPdf: "Compiling PDF...",
      successToast: "Document successfully converted to PDF!",
      errorToast: "An error occurred while processing the file.",
      invalidFileType: "Please upload a valid .docx or .doc file!",
      previewTitle: "Document Preview",
      zoomIn: "Zoom In",
      zoomOut: "Zoom Out",
      removeBtn: "Change File"
    }
  };

  let currentLang = 'uz';

  // --- State Variables ---
  let activeFile = null;
  let activeArrayBuffer = null;
  let zoomLevel = 1.0;
  let isConverting = false;

  // --- DOM Elements ---
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const fileActiveSection = document.getElementById('fileActiveSection');
  const previewContainer = document.getElementById('previewContainer');
  const documentPreview = document.getElementById('documentPreview');
  
  const fileNameDisplay = document.getElementById('fileNameDisplay');
  const fileSizeDisplay = document.getElementById('fileSizeDisplay');
  const removeFileBtn = document.getElementById('removeFileBtn');
  
  const convertBtn = document.getElementById('convertBtn');
  const vectorPrintBtn = document.getElementById('vectorPrintBtn');
  
  const progressCard = document.getElementById('progressCard');
  const progressFill = document.getElementById('progressFill');
  const progressPercent = document.getElementById('progressPercent');
  const progressStatusText = document.getElementById('progressStatusText');
  
  const pageSizeSelect = document.getElementById('pageSizeSelect');
  const orientationSelect = document.getElementById('orientationSelect');
  const marginsSelect = document.getElementById('marginsSelect');
  const qualitySelect = document.getElementById('qualitySelect');
  
  const zoomInBtn = document.getElementById('zoomInBtn');
  const zoomOutBtn = document.getElementById('zoomOutBtn');
  const zoomResetBtn = document.getElementById('zoomResetBtn');
  const themeToggleBtn = document.getElementById('themeToggleBtn');
  const langSelect = document.getElementById('langSelect');
  const toastContainer = document.getElementById('toastContainer');

  // --- Initialization ---
  function init() {
    setupEventListeners();
    setupTheme();
    updateLanguage(currentLang);
  }

  // --- Theme Toggle ---
  function setupTheme() {
    const savedTheme = localStorage.getItem('wtp_theme') || 'dark';
    document.documentElement.setAttribute('data-theme', savedTheme);
    updateThemeIcon(savedTheme);

    themeToggleBtn.addEventListener('click', () => {
      const current = document.documentElement.getAttribute('data-theme') || 'dark';
      const next = current === 'dark' ? 'light' : 'dark';
      document.documentElement.setAttribute('data-theme', next);
      localStorage.setItem('wtp_theme', next);
      updateThemeIcon(next);
    });
  }

  function updateThemeIcon(theme) {
    themeToggleBtn.innerHTML = theme === 'dark' ? '☀️' : '🌙';
    themeToggleBtn.title = theme === 'dark' ? 'Yorug\' rejim' : 'Qorong\'u rejim';
  }

  // --- Language Switching ---
  function updateLanguage(lang) {
    currentLang = lang;
    const t = i18n[lang] || i18n.uz;

    document.querySelectorAll('[data-i18n]').forEach(el => {
      const key = el.getAttribute('data-i18n');
      if (t[key]) {
        el.innerHTML = t[key];
      }
    });

    if (langSelect) langSelect.value = lang;
  }

  // --- Event Listeners ---
  function setupEventListeners() {
    // Dropzone Events
    dropzone.addEventListener('click', () => fileInput.click());

    ['dragenter', 'dragover'].forEach(event => {
      dropzone.addEventListener(event, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(event => {
      dropzone.addEventListener(event, (e) => {
        e.preventDefault();
        e.stopPropagation();
        dropzone.classList.remove('dragover');
      });
    });

    dropzone.addEventListener('drop', (e) => {
      if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
        handleFileSelect(e.dataTransfer.files[0]);
      }
    });

    fileInput.addEventListener('change', (e) => {
      if (e.target.files && e.target.files.length > 0) {
        handleFileSelect(e.target.files[0]);
      }
    });

    removeFileBtn.addEventListener('click', resetFile);

    convertBtn.addEventListener('click', generatePdf);
    vectorPrintBtn.addEventListener('click', handleVectorPrint);

    // Zoom Controls
    zoomInBtn.addEventListener('click', () => {
      zoomLevel = Math.min(zoomLevel + 0.15, 2.0);
      applyZoom();
    });

    zoomOutBtn.addEventListener('click', () => {
      zoomLevel = Math.max(zoomLevel - 0.15, 0.5);
      applyZoom();
    });

    zoomResetBtn.addEventListener('click', () => {
      zoomLevel = 1.0;
      applyZoom();
    });

    if (langSelect) {
      langSelect.addEventListener('change', (e) => {
        updateLanguage(e.target.value);
      });
    }

    // Orientation change updates preview container styling
    orientationSelect.addEventListener('change', () => {
      const isLandscape = orientationSelect.value === 'landscape';
      if (isLandscape) {
        documentPreview.style.maxWidth = '1122px';
        documentPreview.style.minHeight = '794px';
      } else {
        documentPreview.style.maxWidth = '794px';
        documentPreview.style.minHeight = '842px';
      }
    });
  }

  function applyZoom() {
    documentPreview.style.transform = `scale(${zoomLevel})`;
    zoomResetBtn.textContent = `${Math.round(zoomLevel * 100)}%`;
  }

  // --- File Processing ---
  async function handleFileSelect(file) {
    if (!file) return;

    const lowerName = file.name.toLowerCase();
    if (!lowerName.endsWith('.docx') && !lowerName.endsWith('.doc')) {
      showToast(i18n[currentLang].invalidFileType, 'error');
      return;
    }

    activeFile = file;
    fileNameDisplay.textContent = file.name;
    fileSizeDisplay.textContent = formatBytes(file.size);

    dropzone.style.display = 'none';
    fileActiveSection.classList.add('active');
    previewContainer.classList.add('active');

    showProgress(i18n[currentLang].readingFile, 20);

    try {
      const arrayBuffer = await file.arrayBuffer();
      activeArrayBuffer = arrayBuffer;

      showProgress(i18n[currentLang].renderingPages, 50);

      // Render Word to Preview DOM
      await renderWordDocument(arrayBuffer);

      showProgress(i18n[currentLang].ready, 100);
      setTimeout(hideProgress, 800);
    } catch (err) {
      console.error("Faylni o'qishda xatolik:", err);
      showToast(i18n[currentLang].errorToast + ": " + err.message, 'error');
      hideProgress();
    }
  }

  function resetFile() {
    activeFile = null;
    activeArrayBuffer = null;
    fileInput.value = '';
    documentPreview.innerHTML = '';
    
    fileActiveSection.classList.remove('active');
    previewContainer.classList.remove('active');
    dropzone.style.display = 'block';
    hideProgress();
  }

  // --- Document Rendering (docx-preview + mammoth fallback) ---
  async function renderWordDocument(arrayBuffer) {
    documentPreview.innerHTML = '';

    // Check if docx-preview is loaded
    if (window.docx && typeof window.docx.renderAsync === 'function') {
      try {
        const renderOptions = {
          className: 'docx-rendered',
          inWrapper: false,
          ignoreWidth: false,
          ignoreHeight: false,
          ignoreFonts: false,
          breakPages: true,
          ignoreLastRenderedPageBreak: false,
          experimental: true,
          trimXml: true,
          useBase64URL: true,
          renderChanges: false,
          renderHeaders: true,
          renderFooters: true,
          renderFootnotes: true,
          renderEndnotes: true
        };

        await window.docx.renderAsync(arrayBuffer, documentPreview, null, renderOptions);
        return;
      } catch (docxErr) {
        console.warn("docx-preview tahlilida xatolik, mammoth ga o'tilmoqda:", docxErr);
      }
    }

    // Fallback: Mammoth.js
    if (window.mammoth && typeof window.mammoth.convertToHtml === 'function') {
      try {
        const result = await window.mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
        documentPreview.innerHTML = result.value || '<p>Hujjat matni topilmadi.</p>';
        return;
      } catch (mammothErr) {
        console.error("Mammoth xatosi:", mammothErr);
      }
    }

    // If both failed
    documentPreview.innerHTML = `
      <div style="padding: 40px; text-align: center; color: #64748b;">
        <h3>Hujjat strukturasini tahlil qilishda cheklov</h3>
        <p>Hujjat yuklandi, ammo brauzer ko'rinishida to'liq aks ettirilmadi. Siz baribir "Chop etish / Vektor PDF" tugmasi orqali to'liq saqlashingiz mumkin.</p>
      </div>
    `;
  }

  // --- PDF Generation via html2pdf ---
  async function generatePdf() {
    if (!activeFile || isConverting) return;

    if (!window.html2pdf) {
      showToast("PDF kutubxonasi yuklanmadi. Vektor chop etishdan foydalaning.", "error");
      return;
    }

    isConverting = true;
    convertBtn.disabled = true;
    const originalText = convertBtn.innerHTML;
    convertBtn.innerHTML = `<span>⏳</span> ${i18n[currentLang].converting}`;

    showProgress(i18n[currentLang].generatingPdf, 35);

    try {
      // PDF options
      const pageSize = pageSizeSelect.value || 'a4';
      const orientation = orientationSelect.value || 'portrait';
      const marginVal = parseInt(marginsSelect.value, 10) || 10;
      const qualityScale = parseFloat(qualitySelect.value) || 2.0;

      const outputFileName = activeFile.name.replace(/\.[^/.]+$/, "") + ".pdf";

      // Temporary element to format properly without affecting preview zoom
      const clonedElement = documentPreview.cloneNode(true);
      clonedElement.style.transform = 'none';
      clonedElement.style.width = orientation === 'landscape' ? '1120px' : '794px';
      clonedElement.style.maxWidth = 'none';
      clonedElement.style.padding = '20px';
      clonedElement.style.background = '#ffffff';
      clonedElement.style.color = '#000000';
      clonedElement.style.position = 'fixed';
      clonedElement.style.top = '-99999px';
      clonedElement.style.left = '-99999px';
      document.body.appendChild(clonedElement);

      showProgress(i18n[currentLang].generatingPdf, 65);

      const opt = {
        margin: [marginVal, marginVal, marginVal, marginVal],
        filename: outputFileName,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: {
          scale: qualityScale,
          useCORS: true,
          logging: false,
          scrollY: 0
        },
        jsPDF: {
          unit: 'mm',
          format: pageSize,
          orientation: orientation
        },
        pagebreak: { mode: ['avoid-all', 'css', 'legacy'] }
      };

      await window.html2pdf().set(opt).from(clonedElement).save();

      document.body.removeChild(clonedElement);

      showProgress(i18n[currentLang].ready, 100);
      showToast(i18n[currentLang].successToast, 'success');
      setTimeout(hideProgress, 1000);
    } catch (err) {
      console.error("PDF generatsiyasida xatolik:", err);
      showToast(i18n[currentLang].errorToast + ": " + err.message, 'error');
      hideProgress();
    } finally {
      isConverting = false;
      convertBtn.disabled = false;
      convertBtn.innerHTML = originalText;
    }
  }

  // --- Vector Native Print-to-PDF ---
  function handleVectorPrint() {
    if (!activeFile) return;

    const pageSize = pageSizeSelect.value || 'A4';
    const orientation = orientationSelect.value || 'portrait';

    const printWindow = window.open('', '_blank');
    if (!printWindow) {
      showToast("Pop-up oyna bloklandi. Brauzer sozlamalarida ruxsat bering.", "error");
      return;
    }

    const docContent = documentPreview.innerHTML;
    const printDoc = printWindow.document;

    printDoc.open();
    printDoc.write(`
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <title>${activeFile.name.replace(/\.[^/.]+$/, "")}</title>
        <style>
          @page {
            size: ${pageSize} ${orientation};
            margin: 10mm;
          }
          body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif;
            color: #111;
            background: #fff;
            margin: 0;
            padding: 10px;
          }
          * {
            box-sizing: border-box;
          }
          img {
            max-width: 100%;
            height: auto;
          }
          table {
            border-collapse: collapse;
            width: 100%;
          }
          table, th, td {
            border: 1px solid #cbd5e1;
            padding: 6px 10px;
          }
          .docx-rendered {
            width: 100%;
          }
        </style>
      </head>
      <body>
        <div>${docContent}</div>
        <script>
          window.onload = function() {
            setTimeout(function() {
              window.print();
            }, 300);
          };
        </script>
      </body>
      </html>
    `);
    printDoc.close();
  }

  // --- Progress Bar Helpers ---
  function showProgress(text, percent) {
    progressCard.classList.add('active');
    progressStatusText.textContent = text;
    progressFill.style.width = `${percent}%`;
    progressPercent.textContent = `${percent}%`;
  }

  function hideProgress() {
    progressCard.classList.remove('active');
    progressFill.style.width = '0%';
    progressPercent.textContent = '0%';
  }

  // --- Toast Notifications ---
  function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    
    let icon = 'ℹ️';
    if (type === 'success') icon = '✅';
    if (type === 'error') icon = '⚠️';

    toast.innerHTML = `<span>${icon}</span> <span>${escapeHtml(message)}</span>`;
    toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      toast.style.transition = 'all 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 4000);
  }

  // --- Utilities ---
  function formatBytes(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
  }

  function escapeHtml(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  // Run on load
  document.addEventListener('DOMContentLoaded', init);
})();
