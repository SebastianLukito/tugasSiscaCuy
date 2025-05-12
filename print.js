// Helper: setiap kata diawali huruf kapital
function toTitleCase(str) {
    return str
    .toLowerCase()
    .split(' ')
    .filter(w => w)               // buang spasi ekstra
    .map(w => w[0].toUpperCase() + w.slice(1))
    .join(' ');
}

// ---- Rain effect ----
class Rain {
    /**
     * @param {string} canvasId        — id elemen <canvas>
     * @param {number} maxAngle        — sudut maksimum (radian) dari vertikal (default ±0.4 ≈23°)
     * @param {number} windChangeInterval — interval ganti angin (ms), default 5000ms
     */
    constructor(canvasId, maxAngle = 0.4, windChangeInterval = 5000) {
    this.canvas            = document.getElementById(canvasId);
    this.ctx               = this.canvas.getContext('2d');
    this.drops             = [];
    this.animId            = null;
    this.maxAngle          = maxAngle;
    this.windChangeInterval= windChangeInterval;
    this.windAngle         = 0;                // sudut angin saat ini
    this.lastWindChange    = 0;                // timestamp terakhir ganti angin

    this.resize();
    window.addEventListener('resize', () => this.resize());
    }

    resize() {
    this.canvas.width  = window.innerWidth;
    this.canvas.height = window.innerHeight;
    }

    start(numDrops = 200) {
    this.drops = [];
    for (let i = 0; i < numDrops; i++) {
        const length = 10 + Math.random() * 20;
        const speed  = 2  + Math.random() * 10;
        // kita tidak simpan sudut di tiap tetesan, cukup global
        this.drops.push({ x: Math.random() * this.canvas.width,
                        y: Math.random() * this.canvas.height,
                        length, speed });
    }
    // inisialisasi angin pertama
    this.windAngle      = (Math.random() * 2 - 1) * this.maxAngle;
    this.lastWindChange = performance.now();
    this.loop();
    }

    stop() {
    if (this.animId) cancelAnimationFrame(this.animId);
    this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
    }

    loop() {
    const now = performance.now();

    // apakah sudah waktunya ganti arah angin?
    if (now - this.lastWindChange > this.windChangeInterval) {
        this.windAngle      = (Math.random() * 2 - 1) * this.maxAngle;
        this.lastWindChange = now;
    }

    const ctx = this.ctx;
    ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
    ctx.strokeStyle = 'rgba(174,194,224,0.35)';
    ctx.lineWidth   = 1;
    ctx.beginPath();

    for (const drop of this.drops) {
        // pergeseran per frame berdasarkan angin global
        const dx = Math.sin(this.windAngle) * drop.speed;
        const dy = Math.cos(this.windAngle) * drop.speed;

        // gambar garis hujan searah angin
        const x2 = drop.x + Math.sin(this.windAngle) * drop.length;
        const y2 = drop.y + Math.cos(this.windAngle) * drop.length;
        ctx.moveTo(drop.x, drop.y);
        ctx.lineTo(x2, y2);

        // update posisi
        drop.x += dx;
        drop.y += dy;

        // reset saat keluar layar
        if (drop.y > this.canvas.height || drop.x < 0 || drop.x > this.canvas.width) {
        drop.y = -drop.length;
        drop.x = Math.random() * this.canvas.width;
        }
    }

    ctx.stroke();
    this.animId = requestAnimationFrame(() => this.loop());
    }
}

document.addEventListener('DOMContentLoaded', function () {
    // audio background
    const rainAudio      = document.getElementById('rain-audio');
    const audioModal     = document.getElementById('audio-modal');
    const audioConfirm   = document.getElementById('audio-confirm-btn');
    // --- Lightning effect (distant flashes) ---
    const lightningOverlay = document.getElementById('lightning-overlay');

    // 1. Tampilkan modal (CSS .audio-modal sudah display:flex)
    audioModal.style.display = 'flex';

    // 2. Saat user klik tombol: unmute, play, sembunyikan modal
    audioConfirm.addEventListener('click', () => {
        rainAudio.muted = false;           // buka mute
        rainAudio.play().catch(console.error);
        audioModal.style.display = 'none';
    });
        
    // inisialisasi rain effect
    const rain = new Rain('rain-canvas', 0.52 /*≈30°*/, 18000 /*ms*/);
    rain.start(450);

    function triggerFlash() {
        lightningOverlay.classList.add('flash');
        lightningOverlay.addEventListener('animationend', () => {
        lightningOverlay.classList.remove('flash');
        }, { once: true });
    }
    
    function scheduleLightning() {
        // jeda acak antara 5–25 detik
        const delay = Math.random() * 10000 + 500;
        setTimeout(() => {
        // dua kilatan berurutan dengan jarak 100–300ms
        triggerFlash();
        setTimeout(triggerFlash, Math.random() * 200 + 100);
        // jadwalkan storm berikutnya
        scheduleLightning();
        }, delay);
    }
    
    // mulai scheduling
    scheduleLightning();

    // Element references
    const mergeBtn       = document.getElementById('merge-btn');
    const resetBtn       = document.getElementById('reset-btn');
    const dropSection    = document.getElementById('drop-section');
    const fileList       = document.getElementById('file-list');
    const excelInput     = document.getElementById('excelInput');
    const excelPreview   = document.getElementById('excel-preview');
    const tableContainer = document.getElementById('table-container');
    const qrDisplay      = document.getElementById('qr-display');
    const downloadAllBtn = document.getElementById('download-all');
    const sendAllBtn     = document.getElementById('send-all');
    const loadingPopup   = document.getElementById('loading-popup');

    // References for GIF & panel
    const panelContent     = document.querySelector('.panel-content');
    const loadingGif     = document.getElementById('loading-gif');

    // referensi dropdown manual
    const manualToggle   = document.getElementById('manual-toggle');
    const manualContainer= document.getElementById('manual-input-container');
    const manualNameInput= document.getElementById('manual-name');
    const addManualBtn   = document.getElementById('add-manual');

    let manualOpen = false;

    let excelFile = null;
    let excelData = null;
    let individualCertificates = [];

    // Fungsi toggle dengan efek bouncing
    manualToggle.addEventListener('click', () => {
        manualOpen = !manualOpen;
        if (manualOpen) {
        // buka dropdown
        manualContainer.classList.remove('hidden', 'bounce-out');
        manualContainer.classList.add('bounce-in');
        } else {
        // tutup dropdown
        manualContainer.classList.remove('bounce-in');
        manualContainer.classList.add('bounce-out');
        manualContainer.addEventListener('animationend', () => {
            manualContainer.classList.add('hidden');
        }, { once: true });
        }
    });

    // Event tombol "Tambah & Buat" untuk langsung generate satu sertifikat
    addManualBtn.addEventListener('click', async () => {
        const name = manualNameInput.value.trim();
        if (!name) {
        alert('Masukkan nama dulu!');
        return;
        }
        // siapkan data peserta tunggal
        const participant = { name: toTitleCase(name), date: '', email: '' };
        try {
        loadingPopup.style.display = 'flex';
        // load template seperti biasa
        const templateBytes = await fetch('sertifikat.jpg').then(res => {
            if (!res.ok) throw new Error('Template tidak ditemukan');
            return res.arrayBuffer();
        });
        // buat PDF
        const pdfDoc = await PDFLib.PDFDocument.create();
        const font   = await pdfDoc.embedStandardFont(PDFLib.StandardFonts.Helvetica);
        await createCertificatePage(pdfDoc, participant, templateBytes, font);
        const pdfBytes = await pdfDoc.save();
        individualCertificates = [{ name: participant.name, email: '', pdfBytes }];
        renderCertificateList(individualCertificates);
        document.querySelector('.right-panel').classList.add('active');
        hideWithBounce(loadingGif);
        showWithBounce(panelContent);
        } catch (err) {
        console.error(err);
        alert(`Gagal membuat sertifikat manual: ${err.message}`);
        } finally {
        loadingPopup.style.display = 'none';
        }
    });


    // Helper functions for bounce animations
    function showWithBounce(el) {
    el.classList.remove('hidden', 'bounce-out');
    el.classList.add('bounce-in');
    }
    function hideWithBounce(el) {
    el.classList.remove('bounce-in');
    el.classList.add('bounce-out');
    el.addEventListener('animationend', () => {
        el.classList.add('hidden');
    }, { once: true });
    }

    // Initial state: hide panel, show GIF
    panelContent.classList.add('hidden');
    showWithBounce(loadingGif);

    // --- Drag & Drop Excel ---
    dropSection.addEventListener('dragover', e => {
    e.preventDefault();
    dropSection.classList.add('dragover');
    });
    dropSection.addEventListener('dragleave', e => {
    e.preventDefault();
    dropSection.classList.remove('dragover');
    });
    dropSection.addEventListener('drop', e => {
    e.preventDefault();
    dropSection.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length && files[0].name.endsWith('.xlsx')) {
        excelFile = files[0];
        displayFile(files[0]);
        processExcelPreview(files[0]);
        dropSection.querySelector('p').style.display = 'none';
    } else {
        alert('Hanya file Excel (.xlsx) yang diperbolehkan!');
    }
    // stopRainIfReady();
    });
    dropSection.addEventListener('click', () => excelInput.click());
    excelInput.addEventListener('change', e => {
    const file = e.target.files[0];
    if (file && file.name.endsWith('.xlsx')) {
        excelFile = file;
        displayFile(file);
        processExcelPreview(file);
        dropSection.querySelector('p').style.display = 'none';
    } else {
        alert('Format file tidak valid. Harap pilih file .xlsx');
        excelInput.value = '';
    }
    // stopRainIfReady();
    });

    function displayFile(file) {
    fileList.innerHTML = `
        <div class="file-item">
        <img src="excel.png" alt="Excel Icon">
        <span>${file.name}</span>
        </div>
    `;
    }

    resetBtn.addEventListener('click', () => {
        // 1. Hide panel dan reset state
        hideWithBounce(panelContent);
        document.querySelector('.right-panel').classList.remove('active');
        excelFile = null;
        excelData = null;
        individualCertificates = [];
        fileList.innerHTML = '';
        excelPreview.style.display = 'none';
        tableContainer.innerHTML = '';
        excelInput.value = '';
        dropSection.querySelector('p').style.display = 'block';
    
        // 2. Bersihkan semua kelas animasi lama di GIF
        loadingGif.classList.remove('hidden', 'bounce-in', 'bounce-out');
        void loadingGif.offsetWidth; // paksa reflow
    
        // 3. Tampilkan lagi dengan animasi masuk
        showWithBounce(loadingGif);
    
        // 4. Reload page
        window.location.reload();
    });

    // --- Preview Excel hanya Nama & Email ---
    async function processExcelPreview(file) {
    try {
        loadingPopup.style.display = 'flex';
        const data = await processExcel(file);
        excelData = data;

        if (data.length > 0) {
            const displayKeys = ['name', 'email'];
            const labels      = { name: 'Nama', email: 'Email' };

            let html = '<table class="excel-table"><thead><tr>';
            displayKeys.forEach(key => {
                html += `<th>${labels[key]}</th>`;
            });
            html += '</tr></thead><tbody>';

            data.forEach(row => {
                html += '<tr>';
                displayKeys.forEach(key => {
                html += `<td>${row[key] || ''}</td>`;
                });
                html += '</tr>';
            });
            html += '</tbody></table>';

            tableContainer.innerHTML = html;
            excelPreview.style.display = 'block';
            
        }
    } catch (err) {
        console.error('Error previewing Excel:', err);
        alert(`Gagal membaca file Excel: ${err.message}`);
    } finally {
        loadingPopup.style.display = 'none';
    }
    }

    // --- Baca & Parse Excel: nama, tanggal, email ---
    async function processExcel(file) {
        return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => {
            try {
            const arr = new Uint8Array(e.target.result);
            const wb  = XLSX.read(arr, { type: 'array' });
            const ws  = wb.Sheets[wb.SheetNames[0]];
            const rows= XLSX.utils.sheet_to_json(ws, {
                header: 1, defval: '', blankrows: false
            });

            if (rows.length < 2)
                throw new Error('File Excel tidak memiliki data cukup');

            const headers = rows[0].map(h =>
                typeof h === 'string' ? h.trim().toLowerCase() : ''
            );
            const nameIdx  = headers.findIndex(h => h.includes('nama') || h.includes('name'));
            const dateIdx  = headers.findIndex(h => h.includes('tanggal') || h.includes('date'));
            const emailIdx = headers.findIndex(h => h.includes('email'));

            if (nameIdx < 0) throw new Error('Kolom nama tidak ditemukan');

            const participants = rows.slice(1)
                .map(r => {
                    const rawName = r[nameIdx] || '';
                    return {
                        name : toTitleCase(rawName),
                        date : r[dateIdx]  || '',
                        email: emailIdx >= 0 ? r[emailIdx] || '' : ''
                    };
                })
                .filter(p => p.name.trim() !== '');

            if (!participants.length)
                throw new Error('Tidak ada data peserta valid');

            resolve(participants);
            } catch (err) {
            console.error('Error processing Excel:', err);
            reject(new Error('Gagal memproses file Excel. Pastikan kolom nama, tanggal, & email ada.'));
            }
        };
        reader.onerror = () => reject(new Error('Gagal membaca file'));
        reader.readAsArrayBuffer(file);
        });
    }

    // --- Generate PDF per peserta ---
    mergeBtn.addEventListener('click', async () => {
    if (!excelFile) {
        alert('Harap masukkan file Excel terlebih dahulu!');
        return;
    }
    try {
        loadingPopup.style.display = 'flex';

        // Load template sertifikat
        const templateBytes = await fetch('sertifikat.jpg').then(res => {
            if (!res.ok) throw new Error('Template sertifikat tidak ditemukan');
            return res.arrayBuffer();
        });

        const participants = excelData || await processExcel(excelFile);
        if (!participants.length)
            throw new Error('Tidak ada data peserta ditemukan');

        individualCertificates = [];

        for (const p of participants) {
            const pdfDoc = await PDFLib.PDFDocument.create();
            const font   = await pdfDoc.embedStandardFont(PDFLib.StandardFonts.Helvetica);
            await createCertificatePage(pdfDoc, p, templateBytes, font);
            const pdfBytes = await pdfDoc.save();
            individualCertificates.push({ name: p.name, email: p.email, pdfBytes });
        }

        renderCertificateList(individualCertificates);
        document.querySelector('.right-panel').classList.add('active');
        // setelah sertifikat dibuat: toggle GIF & panel
        hideWithBounce(loadingGif);
        showWithBounce(panelContent);
    } catch (err) {
        console.error(err);
        alert(`Gagal membuat sertifikat: ${err.message}`);
    } finally {
        loadingPopup.style.display = 'none';
    }
    });

    // --- Buat halaman sertifikat ---
    async function createCertificatePage(pdfDoc, participant, templateBytes, font) {
    const jpgImage = await pdfDoc.embedJpg(templateBytes);
    const { width, height } = jpgImage.size();
    const page = pdfDoc.addPage([width, height]);

    page.drawImage(jpgImage, { x: 0, y: 0, width, height });

    // Center teks horizontal
    const text     = participant.name;
    const fontSize = 82;
    const textWidth= font.widthOfTextAtSize(text, fontSize);
    const x        = (width - textWidth) / 2;
    const y        = height - 1520;

    page.drawText(text, { x, y, size: fontSize, font, color: PDFLib.rgb(0,0,0) });
    }

    // --- Render list download di panel kanan ---
    function renderCertificateList(certs) {
    qrDisplay.innerHTML = '';
    const ul = document.createElement('ul');
    ul.className = 'certificate-list';

    certs.forEach(({ name, pdfBytes }) => {
        const li = document.createElement('li');
        li.className = 'certificate-item';

        const link = document.createElement('a');
        const blob = new Blob([pdfBytes], { type: 'application/pdf' });
        link.href     = URL.createObjectURL(blob);
        link.download = `sertifikat-${name.toLowerCase().replace(/\s+/g, '-')}.pdf`;

        const img = document.createElement('img');
        img.src       = 'pdf.png';
        img.alt       = 'PDF';
        img.className = 'certificate-icon';

        const span = document.createElement('span');
        span.textContent = `Sertifikat ${name}`;

        link.append(img, span);
        li.appendChild(link);
        ul.appendChild(li);
    });

    qrDisplay.appendChild(ul);
    }

    // --- Batch download all PDFs as ZIP ---
    if (downloadAllBtn) {
    downloadAllBtn.addEventListener('click', async () => {
        if (!individualCertificates.length) {
        alert('Belum ada sertifikat untuk di-download.');
        return;
        }
        const zip = new JSZip();
        individualCertificates.forEach(({ name, pdfBytes }) => {
        const filename = `sertifikat-${name.toLowerCase().replace(/\s+/g, '-')}.pdf`;
        zip.file(filename, pdfBytes);
        });
        try {
        const content = await zip.generateAsync({ type: 'blob' });
        const url     = URL.createObjectURL(content);
        const a       = document.createElement('a');
        a.href        = url;
        a.download    = 'sertifikat-semua.zip';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        } catch (err) {
        console.error('Error membuat ZIP:', err);
        alert('Gagal membuat file ZIP.');
        }
    });
    }

    // --- Helper: convert PDF bytes ke Base64 ---
    function pdfBytesToBase64(bytes) {
    return new Promise(resolve => {
        const blob   = new Blob([bytes], { type: 'application/pdf' });
        const reader = new FileReader();
        reader.onload = () => {
        const base64 = reader.result.split(',')[1];
        resolve(base64);
        };
        reader.readAsDataURL(blob);
    });
    }

    // --- Kirim semua email ---
    

    // --- Hapus tombol download default (jika ada) ---
    const downloadBtn = document.getElementById('download-btn');
    if (downloadBtn) downloadBtn.remove();
});
