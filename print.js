document.addEventListener('DOMContentLoaded', function () {
    // Element references
    const mergeBtn = document.getElementById('merge-btn');
    const resetBtn = document.getElementById('reset-btn');
    const dropSection = document.getElementById('drop-section');
    const fileList = document.getElementById('file-list');
    const excelInput = document.getElementById('excelInput');
    const excelPreview = document.getElementById('excel-preview');
    const tableContainer = document.getElementById('table-container');
    const qrDisplay = document.getElementById('qr-display');
    const pdfPreview = document.getElementById('pdf-preview');
    const downloadBtn = document.getElementById('download-btn');
    const loadingPopup = document.getElementById('loading-popup');
    
    let excelFile = null;
    let excelData = null;
    let pdfBytes = null;
    let currentPdfPage = 1;
    let totalPdfPages = 0;

    // Drag & Drop untuk Excel .xlsx
    dropSection.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropSection.classList.add('dragover');
    });

    dropSection.addEventListener('dragleave', (e) => {
        e.preventDefault();
        dropSection.classList.remove('dragover');
    });

    dropSection.addEventListener('drop', (e) => {
        e.preventDefault();
        dropSection.classList.remove('dragover');
        const files = e.dataTransfer.files;
        
        if(files.length > 0 && files[0].name.endsWith('.xlsx')) {
            excelFile = files[0];
            displayFile(files[0]);
            processExcelPreview(files[0]);
            // Sembunyikan tulisan Drag & Drop
            document.querySelector('#drop-section p').style.display = 'none';
        } else {
            alert('Hanya file Excel (.xlsx) yang diperbolehkan!');
        }
    });

    // Click to select Excel
    dropSection.addEventListener('click', () => excelInput.click());

    excelInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if(file && file.name.endsWith('.xlsx')) {
            excelFile = file;
            displayFile(file);
            processExcelPreview(file);
            // Sembunyikan tulisan Drag & Drop
            document.querySelector('#drop-section p').style.display = 'none';
        } else {
            alert('Format file tidak valid. Harap pilih file .xlsx');
            excelInput.value = '';
        }
    });

    function displayFile(file) {
        fileList.innerHTML = `
            <div class="file-item">
                <img src="excel.png" alt="Excel Icon">
                <span>${file.name}</span>
            </div>
        `;
    }
    
    // Reset button functionality
    resetBtn.addEventListener('click', () => {
        excelFile = null;
        excelData = null;
        pdfBytes = null;
        fileList.innerHTML = '';
        excelPreview.style.display = 'none';
        pdfPreview.innerHTML = '';
        downloadBtn.style.display = 'none';
        excelInput.value = '';
        
        // Tampilkan kembali tulisan Drag & Drop
        document.querySelector('#drop-section p').style.display = 'block';
        // Nonaktifkan panel kanan
        document.querySelector('.right-panel').classList.remove('active');
    });

    // Function to display excel data in a table
    async function processExcelPreview(file) {
        try {
            loadingPopup.style.display = 'flex';
            const data = await processExcel(file);
            excelData = data;
            
            if (data.length > 0) {
                // Create table with the data
                const headers = Object.keys(data[0]);
                
                let tableHTML = '<table class="excel-table"><thead><tr>';
                headers.forEach(header => {
                    tableHTML += `<th>${header}</th>`;
                });
                tableHTML += '</tr></thead><tbody>';
                
                data.forEach(row => {
                    tableHTML += '<tr>';
                    headers.forEach(header => {
                        tableHTML += `<td>${row[header] || ''}</td>`;
                    });
                    tableHTML += '</tr>';
                });
                
                tableHTML += '</tbody></table>';
                
                tableContainer.innerHTML = tableHTML;
                excelPreview.style.display = 'block';
            }
        } catch (error) {
            console.error('Error previewing Excel:', error);
            alert(`Gagal membaca file Excel: ${error.message}`);
        } finally {
            loadingPopup.style.display = 'none';
        }
    }

    // Generate certificates when process button is clicked
    mergeBtn.addEventListener('click', async () => {
        if(!excelFile) {
            alert('Harap masukkan file Excel terlebih dahulu!');
            return;
        }
        
        try {
            loadingPopup.style.display = 'flex';
            
            // Ambil template sertifikat
            const templateResponse = await fetch('sertifikat.jpg');
            if(!templateResponse.ok) throw new Error('Template sertifikat tidak ditemukan');
            const templateBlob = await templateResponse.blob();
            
            // Use already processed data if available
            const participants = excelData || await processExcel(excelFile);
            if(participants.length === 0) throw new Error('Tidak ada data peserta ditemukan');
            
            // Generate PDF
            const pdfDoc = await PDFLib.PDFDocument.create();
            for(const participant of participants) {
                // Get the name from "Nama Lengkap" column
                const nama = participant['Nama Lengkap'] || '';
                if(!nama) continue;
                
                await createCertificatePage(
                    nama, 
                    templateBlob, 
                    pdfDoc
                );
            }
            
            // Save PDF bytes for later use
            pdfBytes = await pdfDoc.save();
            
            // Display PDF preview
            renderPDF(pdfBytes);
            
            // Show download button
            downloadBtn.style.display = 'block';
            
            // Aktifkan panel kanan
            document.querySelector('.right-panel').classList.add('active');
            
        } catch(error) {
            console.error('Error:', error);
            alert(`Gagal membuat sertifikat: ${error.message}`);
        } finally {
            loadingPopup.style.display = 'none';
        }
    });

    // Download button event
    downloadBtn.addEventListener('click', () => {
        if (pdfBytes) {
            downloadPDF(pdfBytes);
        }
    });

    async function processExcel(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, {type: 'array'});
                    
                    // Get first sheet
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    
                    // Convert to JSON with proper format
                    const jsonData = XLSX.utils.sheet_to_json(worksheet, {
                        header: 1,
                        rawNumbers: true,
                        blankrows: false,
                        defval: ""
                    });

                    // Map data starting from second row
                    if (jsonData.length >= 2) {
                        const headers = jsonData[0];
                        const participants = jsonData.slice(1).map(row => {
                            return headers.reduce((obj, header, index) => {
                                obj[header] = row[index] || '';
                                return obj;
                            }, {});
                        });
                        resolve(participants);
                    } else {
                        reject(new Error('File Excel tidak memiliki data yang cukup'));
                    }
                } catch(error) {
                    reject(new Error('Gagal memproses file Excel. Pastikan format file benar!'));
                }
            };
            
            reader.onerror = () => reject(new Error('Gagal membaca file'));
            reader.readAsArrayBuffer(file);
        });
    }

    async function createCertificatePage(nama, templateBlob, pdfDoc) {
        const imgUrl = URL.createObjectURL(templateBlob);
        const img = await new Promise((resolve, reject) => {
            const img = new Image();
            img.onload = () => resolve(img);
            img.onerror = (err) => reject(err);
            img.src = imgUrl;
        });
        
        const canvas = document.createElement('canvas');
        canvas.width = img.width;
        canvas.height = img.height;
        const ctx = canvas.getContext('2d');
        
        // Draw template
        ctx.drawImage(img, 0, 0);
        
        // Singkat nama jika lebih dari 2 kata
        let namaPendek = nama;
        const parts = nama.trim().split(' ');
        if (parts.length > 2) {
            namaPendek = parts[0] + ' ' + parts[1] + ' ' + parts[2][0].toUpperCase() + '.';
        }
        
        // Position for name (adjust as needed based on your template)
        ctx.font = 'bold 82px Arial';
        ctx.fillStyle = '#333333';
        ctx.textAlign = 'center';
        ctx.textBaseline = 'middle';
        
        // Draw name on the certificate (adjust the Y position based on your template)
        ctx.fillText(namaPendek.toUpperCase(), canvas.width/2, 1490);
        
        // Convert to PNG and add to PDF
        const pngImage = await pdfDoc.embedPng(canvas.toDataURL());
        const page = pdfDoc.addPage([img.width, img.height]);
        page.drawImage(pngImage, {
            x: 0,
            y: 0,
            width: img.width,
            height: img.height
        });
        
        URL.revokeObjectURL(imgUrl);
    }

    async function renderPDF(pdfBytes) {
        pdfPreview.innerHTML = '';
        
        // Create PDF navigation controls
        const navDiv = document.createElement('div');
        navDiv.className = 'pdf-navigation';
        navDiv.innerHTML = `
            <button id="prev-page">← Sebelumnya</button>
            <span id="page-indicator">Halaman 1 dari 1</span>
            <button id="next-page">Selanjutnya →</button>
        `;
        pdfPreview.appendChild(navDiv);
        
        const pageIndicator = document.getElementById('page-indicator');
        const prevBtn = document.getElementById('prev-page');
        const nextBtn = document.getElementById('next-page');
        
        // Create canvas for PDF display
        const canvas = document.createElement('canvas');
        pdfPreview.appendChild(canvas);
        
        // Load the PDF
        const pdf = await pdfjsLib.getDocument({data: pdfBytes}).promise;
        totalPdfPages = pdf.numPages;
        currentPdfPage = 1;
        
        // Update page indicator
        pageIndicator.textContent = `Halaman ${currentPdfPage} dari ${totalPdfPages}`;
        
        // Render the first page
        const page = await pdf.getPage(currentPdfPage);
        const viewport = page.getViewport({scale: 1.0});
        
        canvas.height = viewport.height;
        canvas.width = viewport.width;
        
        const renderContext = {
            canvasContext: canvas.getContext('2d'),
            viewport: viewport
        };
        
        await page.render(renderContext).promise;
        
        // Previous page button
        prevBtn.addEventListener('click', async () => {
            if (currentPdfPage > 1) {
                currentPdfPage--;
                const page = await pdf.getPage(currentPdfPage);
                const viewport = page.getViewport({scale: 1.0});
                
                canvas.height = viewport.height;
                canvas.width = viewport.width;
                
                const renderContext = {
                    canvasContext: canvas.getContext('2d'),
                    viewport: viewport
                };
                
                await page.render(renderContext).promise;
                pageIndicator.textContent = `Halaman ${currentPdfPage} dari ${totalPdfPages}`;
            }
        });
        
        // Next page button
        nextBtn.addEventListener('click', async () => {
            if (currentPdfPage < totalPdfPages) {
                currentPdfPage++;
                const page = await pdf.getPage(currentPdfPage);
                const viewport = page.getViewport({scale: 1.0});
                
                canvas.height = viewport.height;
                canvas.width = viewport.width;
                
                const renderContext = {
                    canvasContext: canvas.getContext('2d'),
                    viewport: viewport
                };
                
                await page.render(renderContext).promise;
                pageIndicator.textContent = `Halaman ${currentPdfPage} dari ${totalPdfPages}`;
            }
        });
    }

    function downloadPDF(pdfBytes) {
        const blob = new Blob([pdfBytes], {type: 'application/pdf'});
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `Sertifikat_${new Date().toISOString().slice(0,10)}.pdf`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(() => URL.revokeObjectURL(url), 100);
    }
});