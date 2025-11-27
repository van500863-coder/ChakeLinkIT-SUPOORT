document.addEventListener('DOMContentLoaded', () => {
    // --- ELEMENT SELECTORS ---
    const loader = document.getElementById('loader');
    const tableBody = document.getElementById('table-body');
    const statusMessage = document.getElementById('status-message');
    const dateTimeDisplay = document.getElementById('datetime-display');
    const clearAllBtn = document.getElementById('clear-all-btn');
    const daySelector = document.getElementById('day-selector');
    const dynamicDayHeader = document.getElementById('dynamic-day-header');
    
    // ចាប់យក Buttons Download ថ្មី
    const downloadReportBtn = document.getElementById('download-report-btn');
    const downloadImageBtn = document.getElementById('download-image-btn');

    // --- CONFIGURATION ---
    const SHEET_1_ID = '1QOKQyFDTpJAM7j5EY6RpmahstNxvwe4C4FRGC4fw5DE';
    const SHEET_1_NAME = 'DATA';
    const SHEET_2_ID = '1eRyPoifzyvB4oBmruNyXcoKMKPRqjk6xDD6-bPNW6pc';
    const SHEET_2_NAME = 'DIList';
    const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbz0mZp-C1kg58iFLkgVo9hB_2uI5Wx-OVwyoik844LNThDmNb7Gw3hC-n8F1zrXImqO0g/exec';

    const URL_SHEET_1 = `https://docs.google.com/spreadsheets/d/${SHEET_1_ID}/gviz/tq?tqx=out:json&sheet=${SHEET_1_NAME}&range=B3:C31`;
    const URL_SHEET_2 = `https://docs.google.com/spreadsheets/d/${SHEET_2_ID}/gviz/tq?tqx=out:json&sheet=${SHEET_2_NAME}`;

    let allUsersData = [];
    const khmerDays = ["ថ្ងៃអាទិត្យ", "ថ្ងៃចន្ទ", "ថ្ងៃអង្គារ", "ថ្ងៃពុធ", "ថ្ងៃព្រហស្បតិ៍", "ថ្ងៃសុក្រ", "ថ្ងៃសៅរ៍"];

    // --- UTILITY FUNCTIONS ---
    
    function showStatus(message, type) {
        statusMessage.innerHTML = message;
        statusMessage.className = type;
        if (window.statusTimeout) clearTimeout(window.statusTimeout);
        window.statusTimeout = setTimeout(() => { statusMessage.innerHTML = ''; }, 5000);
    }

    function updateDateTime() {
        const now = new Date();
        const days = ["អាទិត្យ", "ចន្ទ", "អង្គារ", "ពុធ", "ព្រហស្បតិ៍", "សុក្រ", "សៅរ៍"];
        const months = ["មករា", "កុម្ភៈ", "មីនា", "មេសា", "ឧសភា", "មិថុនា", "កក្កដា", "សីហា", "កញ្ញា", "តុលា", "វិច្ឆិកា", "ធ្នូ"];
        const dayName = days[now.getDay()];
        const day = now.getDate();
        const monthName = months[now.getMonth()];
        const year = now.getFullYear();
        const timeString = now.toLocaleTimeString('en-US', { hour12: true });
        dateTimeDisplay.innerHTML = `<i class="fa-regular fa-clock"></i> ថ្ងៃ${dayName} ទី${day} ខែ${monthName} ឆ្នាំ${year} | ${timeString}`;
    }

    // --- DATA FETCHING & DISPLAY ---

    async function fetchData() {
        loader.style.display = 'block';
        tableBody.innerHTML = '';
        allUsersData = [];
        try {
            const [res1, res2] = await Promise.all([ fetch(URL_SHEET_1), fetch(URL_SHEET_2) ]);
            const [text1, text2] = await Promise.all([ res1.text(), res2.text() ]);
            
            const data1 = JSON.parse(text1.substring(47, text1.length - 2));
            const namesAndLinks = (data1.table.rows || []).map(row => ({
                name: row.c[0]?.v || '', 
                linkCount: row.c[1]?.v || 0
            })).filter(item => item.name && item.name.trim() !== '');

            const data2 = JSON.parse(text2.substring(47, text2.length - 2));
            const schedules = (data2.table.rows || []).slice(8).map(row => ({
                name: row.c[11]?.v || '',
                schedule: { 'ថ្ងៃចន្ទ': row.c[28]?.v, 'ថ្ងៃអង្គារ': row.c[29]?.v, 'ថ្ងៃពុធ': row.c[30]?.v, 'ថ្ងៃព្រហស្បតិ៍': row.c[31]?.v, 'ថ្ងៃសុក្រ': row.c[32]?.v, 'ថ្ងៃសៅរ៍': row.c[33]?.v, 'ថ្ងៃអាទិត្យ': row.c[34]?.v }
            })).filter(item => item.name && item.name.trim() !== '');

            const mergedData = namesAndLinks.map(person => {
                const personNameLower = person.name.trim().toLowerCase();
                const foundSchedule = schedules.find(s => s.name.trim().toLowerCase() === personNameLower);
                return { ...person, schedule: foundSchedule ? foundSchedule.schedule : {} };
            });

            allUsersData = mergedData;
            displayDataForSelectedDay(); 
        } catch (error) {
            console.error('Error fetching data:', error);
            showStatus('Error: មិនអាចទាញទិន្នន័យបានទេ។ សូមពិនិត្យមើល Internet របស់អ្នក។', 'error');
        } finally {
            loader.style.display = 'none';
        }
    }
    
    function displayDataForSelectedDay() {
        if (allUsersData.length === 0) {
            tableBody.innerHTML = '<tr><td colspan="4" style="text-align:center; padding: 20px;">មិនមានទិន្នន័យ។</td></tr>';
            return;
        }
        const selectedDayIndex = parseInt(daySelector.value, 10);
        const selectedDayName = khmerDays[selectedDayIndex];
        dynamicDayHeader.textContent = `វេន (${selectedDayName})`;

        tableBody.innerHTML = allUsersData.map((item, index) => {
            const schedule = item.schedule || {};
            let rawSchedule = schedule[selectedDayName];
            
            // --- CONFLICT LOGIC (Fix: "-" and empty strings become "ប៉ះម៉ោងរៀន") ---
            let displayText = rawSchedule;
            let isConflict = false;

            const strSchedule = rawSchedule ? rawSchedule.toString().trim() : '';

            if (strSchedule === '' || strSchedule === '-' || strSchedule === 'null' || strSchedule.includes('ប៉ះម៉ោងរៀន')) {
                displayText = 'ប៉ះម៉ោងរៀន';
                isConflict = true;
            }

            const scheduleHtml = isConflict 
                ? `<span class="conflict-badge"><i class="fa-solid fa-circle-exclamation"></i> ${displayText}</span>` 
                : (displayText ? displayText : '<span style="opacity:0.5">-</span>');
            // ------------------------------------------------------------------------

            return `
                <tr>
                    <td><div style="font-weight:600;">${item.name}</div></td>
                    <td>
                        <div class="number-input-container">
                            <button class="number-input-btn" onclick="decrementValue(${index})"><i class="fa-solid fa-minus"></i></button>
                            <input type="number" class="link-input" id="link-input-${index}" value="${item.linkCount}" min="0" oninput="markAsDirty(${index})">
                            <button class="number-input-btn" onclick="incrementValue(${index})"><i class="fa-solid fa-plus"></i></button>
                        </div>
                    </td>
                    <td>
                        <div class="save-container">
                            <button class="save-btn" onclick="saveData(this, '${item.name}', ${index})"><i class="fa-regular fa-floppy-disk"></i> Save</button>
                            <div class="status-dot" id="status-dot-${index}"></div>
                        </div>
                    </td>
                    <td>${scheduleHtml}</td>
                </tr>
            `;
        }).join('');
    }

    function setDefaultDay() {
        const todayIndex = new Date().getDay();
        daySelector.value = todayIndex;
    }
    
    // --- TABLE INTERACTION WINDOW FUNCTIONS ---

    window.markAsDirty = (index) => {
        const dot = document.getElementById(`status-dot-${index}`);
        if (dot) dot.classList.add('visible');
    };

    window.incrementValue = (index) => {
        const input = document.getElementById(`link-input-${index}`);
        let value = parseInt(input.value, 10);
        value = isNaN(value) ? 0 : value;
        value++;
        input.value = value;
        markAsDirty(index);
    };

    window.decrementValue = (index) => {
        const input = document.getElementById(`link-input-${index}`);
        let value = parseInt(input.value, 10);
        value = isNaN(value) ? 0 : value;
        if (value > 0) {
            value--;
            input.value = value;
            markAsDirty(index);
        }
    };

    // --- SAVE & CLEAR LOGIC ---
    
    window.saveData = async (button, name, index) => {
        const input = document.getElementById(`link-input-${index}`);
        const dot = document.getElementById(`status-dot-${index}`);
        const originalText = button.innerHTML;
        
        if (dot) dot.classList.remove('visible');
        
        button.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Saving...';
        button.disabled = true;
        
        const result = await updateSingleUser(name, input.value);
        
        if (result.success) { 
            showStatus(`<i class="fa-solid fa-check-circle"></i> ${result.message}`, 'success'); 
        } else { 
            showStatus(`<i class="fa-solid fa-triangle-exclamation"></i> Error: ${result.message}`, 'error');
            markAsDirty(index);
        }
        
        button.innerHTML = originalText;
        button.disabled = false;
    };

    async function updateSingleUser(name, linkCount) {
        try {
            const response = await fetch(APPS_SCRIPT_URL, {
                method: 'POST', mode: 'cors',
                body: JSON.stringify({ name: name, linkCount: linkCount })
            });
            const result = await response.json();
            if (result.status === 'success') return { success: true, message: result.message };
            else throw new Error(result.message);
        } catch (error) {
            console.error('Error saving data for', name, ':', error);
            return { success: false, message: error.message };
        }
    }

    // --- DOWNLOAD FUNCTIONS (FIXED Font Rendering) ---

    async function downloadPDF() {
        const { jsPDF } = window.jspdf;
        const reportContent = document.getElementById('report-content');
        
        if (!reportContent || !downloadReportBtn) return;

        const originalText = downloadReportBtn.innerHTML;
        downloadReportBtn.disabled = true;
        downloadReportBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Processing...';
        showStatus('កំពុងបង្កើត PDF...', 'success');

        try {
            const canvas = await html2canvas(reportContent, {
                scale: 2,
                backgroundColor: '#1e293b',
                logging: false,
                useCORS: true,
                // FIX FONT ISSUE: Remove letter-spacing only during capture
                onclone: (clonedDoc) => {
                    const ths = clonedDoc.querySelectorAll('#data-table th');
                    ths.forEach(th => {
                        th.style.letterSpacing = 'normal'; 
                    });
                }
            });

            const imgData = canvas.toDataURL('image/png');
            const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
            const pdfWidth = pdf.internal.pageSize.getWidth();
            const imgProps = pdf.getImageProperties(imgData);
            const imgHeight = (imgProps.height * pdfWidth) / imgProps.width;

            pdf.setFillColor(30, 41, 59);
            pdf.rect(0, 0, pdfWidth, 20, "F");
            pdf.setFontSize(16);
            pdf.setTextColor(255, 255, 255);
            pdf.text('IT-SUPPORT REPORT', pdfWidth / 2, 13, { align: 'center' });

            pdf.addImage(imgData, 'PNG', 0, 25, pdfWidth, imgHeight);

            const today = new Date().toISOString().slice(0, 10);
            pdf.save(`Report-IT-Support-${today}.pdf`);
            showStatus('បាន Download PDF ជោគជ័យ!', 'success');

        } catch (error) {
            console.error("PDF Error:", error);
            alert("បរាជ័យ៖ " + error.message);
        } finally {
            downloadReportBtn.disabled = false;
            downloadReportBtn.innerHTML = originalText;
        }
    }

    async function downloadImage() {
        const reportContent = document.getElementById('report-content');
        
        if (!reportContent || !downloadImageBtn) return;

        const originalText = downloadImageBtn.innerHTML;
        downloadImageBtn.disabled = true;
        downloadImageBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Processing...';
        showStatus('កំពុងបង្កើតរូបភាព...', 'success');

        try {
            const canvas = await html2canvas(reportContent, {
                scale: 2,
                backgroundColor: '#1e293b',
                logging: false,
                useCORS: true,
                // FIX FONT ISSUE: Remove letter-spacing only during capture
                onclone: (clonedDoc) => {
                    const ths = clonedDoc.querySelectorAll('#data-table th');
                    ths.forEach(th => {
                        th.style.letterSpacing = 'normal';
                    });
                }
            });

            const link = document.createElement('a');
            link.download = `Report-IT-Support-${new Date().toISOString().slice(0, 10)}.png`;
            link.href = canvas.toDataURL('image/png');
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            
            showStatus('បាន Download រូបភាពជោគជ័យ!', 'success');

        } catch (error) {
            console.error("Image Error:", error);
            alert("បរាជ័យ៖ " + error.message);
        } finally {
            downloadImageBtn.disabled = false;
            downloadImageBtn.innerHTML = originalText;
        }
    }

    // --- EVENT LISTENERS ---
    
    daySelector.addEventListener('change', displayDataForSelectedDay);

    clearAllBtn.addEventListener('click', async () => {
        if (!confirm('តើអ្នកពិតជាចង់កំណត់ចំនួន Link ទាំងអស់ទៅ 0 មែនទេ?')) return;
        clearAllBtn.disabled = true;
        clearAllBtn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> Clearing...';
        showStatus('កំពុងលុបទិន្នន័យ...', 'success');
        
        const updatePromises = allUsersData.map(user => updateSingleUser(user.name, 0));
        
        try {
            await Promise.all(updatePromises);
            showStatus('បានលុបចំនួន Link ទាំងអស់ដោយជោគជ័យ!', 'success');
            await fetchData();
        } catch (error) {
            showStatus('Error: មានបញ្ហាក្នុងការលុបទិន្នន័យ។', 'error');
        } finally {
            clearAllBtn.disabled = false;
            clearAllBtn.innerHTML = '<i class="fa-solid fa-trash-can"></i> Clear All';
        }
    });

    // ចង Event របស់ Download Buttons នៅក្នុង DOMContentLoaded ដើម្បីការពារ Scope Issue (ដំណោះស្រាយចុងក្រោយ)
    if (downloadReportBtn) {
        downloadReportBtn.addEventListener('click', downloadPDF);
    }
    if (downloadImageBtn) {
        downloadImageBtn.addEventListener('click', downloadImage);
    }

    // --- Initial Load ---
    setDefaultDay();
    updateDateTime();
    setInterval(updateDateTime, 1000);
    fetchData();

});
