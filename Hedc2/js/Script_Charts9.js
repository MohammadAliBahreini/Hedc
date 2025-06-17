// ====================================================================================================
// تعریف متغیرهای سراسری
// این متغیرها نیازی به دسترسی مستقیم به DOM در زمان تعریف ندارند و در طول برنامه استفاده می‌شوند.
// ====================================================================================================
let workbook; // برای نگهداری شیء ورک‌بوک اکسل پس از بارگذاری فایل
let parsedData = []; // برای نگهداری آرایه‌ای از اطلاعات پردازش شده مشترکین
let currentCharts = []; // آرایه‌ای برای نگهداری نمونه‌های نمودار Chart.js (برای مدیریت و پاکسازی)

// ====================================================================================================
// توابع کمکی
// این توابع عملیات‌های خاصی را انجام می‌دهند که در بخش‌های مختلف برنامه فراخوانی می‌شوند.
// ====================================================================================================

/**
 * تابع نمایش و به‌روزرسانی نوار پیشرفت
 * @param {number} percent - درصد پیشرفت (0 تا 100)
 * @param {string} label - متن نمایشی نوار پیشرفت
 */
function showProgress(percent, label = 'در حال بارگذاری...') {
    const progressBar = document.getElementById('progress-bar');
    const progressLabel = document.getElementById('progress-label');
    const progressContainer = document.getElementById('progress-container');

    if (progressBar && progressLabel && progressContainer) {
        progressBar.style.width = percent + '%';
        progressBar.setAttribute('aria-valuenow', percent);
        progressLabel.textContent = label;
        progressContainer.style.display = 'block'; // نمایش نوار پیشرفت
    }
}

/**
 * تابع مخفی کردن نوار پیشرفت
 */
function hideProgress() {
    const progressContainer = document.getElementById('progress-container');
    if (progressContainer) {
        progressContainer.style.display = 'none';
    }
}

/**
 * تابع پاکسازی نمودارهای موجود
 * این تابع، هر نمونه نمودار Chart.js را از حافظه و DOM حذف می‌کند.
 */
function destroyCharts() {
    currentCharts.forEach(chart => chart.destroy());
    currentCharts = []; // آرایه نمودارهای موجود را خالی می‌کند
    const chartsContainer = document.getElementById('chartsContainer');
    if (chartsContainer) {
        // حذف تمام فرزندان به جز h2 و p#noChartsMessage
        Array.from(chartsContainer.children).forEach(child => {
            if (child.tagName !== 'H2' && child.id !== 'noChartsMessage') {
                child.remove();
            }
        });
    }
}

/**
 * تابع نمایش پیام عدم وجود نمودار
 */
function showNoChartsMessage() {
    const noChartsMessage = document.getElementById('noChartsMessage');
    if (noChartsMessage) {
        noChartsMessage.style.display = 'block';
    }
}

/**
 * تابع پنهان کردن پیام عدم وجود نمودار
 */
function hideNoChartsMessage() {
    const noChartsMessage = document.getElementById('noChartsMessage');
    if (noChartsMessage) {
        noChartsMessage.style.display = 'none';
    }
}

// ====================================================================================================
// توابع اصلی برنامه
// این توابع جریان اصلی برنامه را مدیریت می‌کنند و با DOM تعامل دارند.
// ====================================================================================================

document.addEventListener('DOMContentLoaded', () => {
    // ================================================================================================
    // ارجاع به عناصر DOM
    // این متغیرها پس از بارگذاری کامل DOM مقداردهی می‌شوند.
    // ================================================================================================
    const fileInput = document.getElementById('excelFile');
    const fileNameDisplay = document.getElementById('fileNameDisplay');
    const sheetSelect = document.getElementById('sheetSelect');
    const processBtn = document.getElementById('processDataBtn');
    const resetAppBtn = document.getElementById('resetAppBtn');
    const resultsTableBody = document.querySelector('#resultsTable tbody');
    const chartsContainer = document.getElementById('chartsContainer');
    const exportExcelBtn = document.getElementById('exportExcelBtn');
    const exportPdfBtn = document.getElementById('exportPdfBtn');
    const resultsContainer = document.getElementById('resultsContainer'); // تعریف در اینجا
    const noChartsMessage = document.getElementById('noChartsMessage');

    const chkMorning = document.getElementById('chkMorning');
    const txtMorning = document.getElementById('txtMorning');
    const chkEvening = document.getElementById('chkEvening');
    const txtEvening = document.getElementById('txtEvening');
    const chkMaxReduction = document.getElementById('chkMaxReduction');
    const txtMaxReduction = document.getElementById('txtMaxReduction');
    const chkMinEveningLoad = document.getElementById('chkMinEveningLoad');
    const txtMinEveningLoad = document.getElementById('txtMinEveningLoad');


    // ================================================================================================
    // شنونده‌های رویداد (Event Listeners)
    // این بخش، رویدادهای مختلف را مدیریت می‌کند و توابع مربوطه را فراخوانی می‌کند.
    // ================================================================================================

    /**
     * رویداد تغییر فایل ورودی (Excel)
     * فایل انتخاب شده را می‌خواند و برگه‌های آن را در دراپ‌داون نمایش می‌دهد.
     */
    fileInput.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (!file) {
            fileNameDisplay.textContent = 'فایل انتخاب نشده...';
            sheetSelect.innerHTML = '<option value="">- Sheet1 -</option>';
            sheetSelect.disabled = true;
            processBtn.disabled = true;
            exportExcelBtn.disabled = true;
            exportPdfBtn.disabled = true;
            if (resultsTableBody) resultsTableBody.innerHTML = '';
            // اطمینان از وجود resultsContainer قبل از دسترسی به style
            if (resultsContainer) resultsContainer.style.display = 'none';
            destroyCharts();
            showNoChartsMessage();
            hideProgress();
            return;
        }

        fileNameDisplay.textContent = file.name;
        showProgress(10, 'در حال خواندن فایل...');

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                workbook = XLSX.read(data, { type: 'array' });
                updateSheetSelect(workbook.SheetNames);
                sheetSelect.disabled = false;
                processBtn.disabled = false;
                hideProgress();
            } catch (error) {
                Swal.fire('خطا', 'خطا در خواندن فایل اکسل. لطفاً از فرمت صحیح (xlsx/xls) مطمئن شوید.', 'error');
                console.error("Error reading Excel file:", error);
                fileInput.value = ''; // پاک کردن فایل انتخاب شده
                fileNameDisplay.textContent = 'فایل انتخاب نشده...';
                sheetSelect.innerHTML = '<option value="">- Sheet1 -</option>';
                sheetSelect.disabled = true;
                processBtn.disabled = true;
                hideProgress();
            }
        };
        reader.onerror = (error) => {
            Swal.fire('خطا', 'خطا در بارگذاری فایل.', 'error');
            console.error("FileReader error:", error);
            hideProgress();
        };
        reader.readAsArrayBuffer(file);
    });

    /**
     * رویداد کلیک دکمه "پردازش داده‌ها"
     * داده‌ها را از شیت انتخاب شده می‌خواند، پردازش می‌کند و نتایج را نمایش می‌دهد.
     */
    processBtn.addEventListener('click', () => {
        if (!workbook) {
            Swal.fire('اخطار', 'لطفاً ابتدا یک فایل اکسل انتخاب کنید.', 'warning');
            return;
        }

        const selectedSheetName = sheetSelect.value || sheetSelect.options[0].value;
        if (!selectedSheetName) {
            Swal.fire('اخطار', 'برگه‌ای برای پردازش یافت نشد.', 'warning');
            return;
        }

        showProgress(20, 'در حال پردازش داده‌ها...');
        setTimeout(() => { // استفاده از setTimeout برای نمایش بهتر نوار پیشرفت
            try {
                const worksheet = workbook.Sheets[selectedSheetName];
                // raw: false برای حفظ فرمت تاریخ و غیره
                // header: 1 برای خواندن ردیف اول به عنوان هدر
                // defval: null برای اینکه خانه‌های خالی null شوند
                const jsonOptions = { header: 1, raw: false, defval: null };
                const jsonData = XLSX.utils.sheet_to_json(worksheet, jsonOptions);

                if (!jsonData || jsonData.length < 2) { // حداقل یک ردیف هدر و یک ردیف داده
                    Swal.fire('اخطار', 'برگه انتخاب شده فاقد داده معتبر است. (حداقل یک ردیف هدر و یک ردیف داده نیاز است).', 'warning');
                    hideProgress();
                    return;
                }

                parsedData = processExcelData(jsonData);
                if (parsedData.length === 0) {
                    Swal.fire('اخطار', 'هیچ داده مشترکی برای پردازش یافت نشد. لطفاً ساختار فایل اکسل را بررسی کنید. (مطمئن شوید ستون‌های ID/Customer id، نام مشترک/Customer name و ۹۶ ستون ساعتی مانند "00:00 to 00:15 [KW]" و همچنین "Contracted demand" برای فیلتر "حداکثر درصد کاهش" وجود دارند).', 'warning');
                    hideProgress();
                    return;
                }

                displayResultsTable(parsedData);
                destroyCharts(); // پاکسازی نمودارهای قبلی قبل از رسم جدید
                hideNoChartsMessage();
                renderCharts(parsedData);

                exportExcelBtn.disabled = false;
                exportPdfBtn.disabled = false;
                // اطمینان از وجود resultsContainer قبل از دسترسی به style
                if (resultsContainer) resultsContainer.style.display = 'block';
                hideProgress();
                Swal.fire('موفقیت', 'داده‌ها با موفقیت پردازش شدند.', 'success');

            } catch (error) {
                Swal.fire('خطا', 'خطا در پردازش داده‌ها. لطفاً ساختار فایل اکسل را بررسی کنید و از فرمت صحیح اطمینان حاصل کنید.', 'error');
                console.error("Error processing data:", error);
                hideProgress();
            }
        }, 100); // تأخیر کم برای نمایش نوار پیشرفت
    });

    /**
     * رویداد کلیک دکمه "شروع مجدد"
     * برنامه را به حالت اولیه بازمی‌گرداند.
     */
    resetAppBtn.addEventListener('click', () => {
        workbook = null;
        parsedData = [];
        destroyCharts(); // پاکسازی نمودارها

        if (fileInput) {
            fileInput.value = '';
            if (fileNameDisplay) fileNameDisplay.textContent = 'فایل انتخاب نشده...';
        }
        if (sheetSelect) {
            sheetSelect.innerHTML = '<option value="">- Sheet1 -</option>';
            sheetSelect.disabled = true;
        }
        if (processBtn) processBtn.disabled = true;
        if (exportExcelBtn) exportExcelBtn.disabled = true;
        if (exportPdfBtn) exportPdfBtn.disabled = true;
        if (resultsTableBody) resultsTableBody.innerHTML = '';
        // اطمینان از وجود resultsContainer قبل از دسترسی به style
        if (resultsContainer) resultsContainer.style.display = 'none'; // مخفی کردن جدول نتایج
        showNoChartsMessage(); // نمایش پیام عدم وجود نمودار

        // بازگرداندن تنظیمات به حالت پیش‌فرض
        if (chkMorning) chkMorning.checked = false;
        if (txtMorning) txtMorning.disabled = true;
        if (chkEvening) chkEvening.checked = false;
        if (txtEvening) txtEvening.disabled = true;
        if (chkMaxReduction) chkMaxReduction.checked = false;
        if (txtMaxReduction) {
            txtMaxReduction.disabled = true;
            txtMaxReduction.value = "60"; // بازگرداندن به مقدار پیش‌فرض
        }
        if (chkMinEveningLoad) chkMinEveningLoad.checked = false;
        if (txtMinEveningLoad) {
            txtMinEveningLoad.disabled = true;
            txtMinEveningLoad.value = "15"; // بازگرداندن به مقدار پیش‌فرض
        }

        hideProgress();
        Swal.fire('با موفقیت', 'برنامه به حالت اولیه بازگشت.', 'success');
    });

    /**
     * رویداد تغییر وضعیت چک‌باکس بار صبح (chkMorning)
     * فیلد ورودی بار صبح را فعال یا غیرفعال می‌کند.
     */
    if (chkMorning && txtMorning) {
        chkMorning.addEventListener('change', () => {
            txtMorning.disabled = !chkMorning.checked;
        });
    }

    /**
     * رویداد تغییر وضعیت چک‌باکس بار عصر (chkEvening)
     * فیلد ورودی بار عصر را فعال یا غیرفعال می‌کند.
     */
    if (chkEvening && txtEvening) {
        chkEvening.addEventListener('change', () => {
            txtEvening.disabled = !chkEvening.checked;
        });
    }

    /**
     * رویداد تغییر وضعیت چک‌باکس حداکثر درصد کاهش (chkMaxReduction)
     * فیلد ورودی حداکثر درصد کاهش را فعال یا غیرفعال می‌کند.
     */
    if (chkMaxReduction && txtMaxReduction) {
        chkMaxReduction.addEventListener('change', () => {
            txtMaxReduction.disabled = !chkMaxReduction.checked;
        });
    }

    /**
     * رویداد تغییر وضعیت چک‌باکس حذف اگر بار عصر کمتر از ۱۵KW (chkMinEveningLoad)
     * فیلد ورودی مربوطه را فعال یا غیرفعال می‌کند.
     */
    if (chkMinEveningLoad && txtMinEveningLoad) {
        chkMinEveningLoad.addEventListener('change', () => {
            txtMinEveningLoad.disabled = !chkMinEveningLoad.checked;
        });
    }

    /**
     * رویداد کلیک دکمه "خروجی اکسل"
     * داده‌های پردازش شده را به فرمت اکسل خروجی می‌گیرد.
     */
    exportExcelBtn.addEventListener('click', () => {
        if (parsedData.length === 0) {
            Swal.fire('اخطار', 'هیچ داده‌ای برای خروجی اکسل وجود ندارد.', 'warning');
            return;
        }

        showProgress(10, 'در حال آماده‌سازی خروجی اکسل...');
        setTimeout(() => { // تأخیر برای نمایش نوار پیشرفت
            try {
                const ws_data = [
                    ["ردیف", "نام مشترک", "شناسه", "مجموع مصرف (KWH)", "پیک مصرف (KW)", "زمان پیک", "بار صبح (KW)", "بار عصر (KW)", "میزان کاهش (KW)", "درصد کاهش (%)"]
                ];

                parsedData.forEach((row, index) => {
                    ws_data.push([
                        index + 1,
                        row.name,
                        row.id,
                        row.totalConsumption,
                        row.peakConsumption,
                        row.peakTime,
                        row.morningLoad, // اضافه شدن بار صبح
                        row.eveningLoad,
                        row.reductionAmount,
                        row.reductionPercent
                    ]);
                });

                const ws = XLSX.utils.aoa_to_sheet(ws_data);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "گزارش تحلیل مصرف");

                XLSX.writeFile(wb, "گزارش_تحلیل_مصرف_برق.xlsx");
                hideProgress();
                Swal.fire('موفقیت', 'فایل اکسل با موفقیت ایجاد شد.', 'success');
            } catch (error) {
                Swal.fire('خطا', 'خطا در ایجاد فایل اکسل.', 'error');
                console.error("Error exporting Excel:", error);
                hideProgress();
            }
        }, 100);
    });

    /**
     * رویداد کلیک دکمه "خروجی PDF پروفیل‌ها"
     * نمودارهای پروفیل را به PDF خروجی می‌گیرد.
     */
    exportPdfBtn.addEventListener('click', async () => {
        if (currentCharts.length === 0) {
            Swal.fire('اخطار', 'هیچ نموداری برای خروجی PDF وجود ندارد.', 'warning');
            return;
        }

        showProgress(10, 'در حال آماده‌سازی خروجی PDF...');
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF('p', 'mm', 'a4'); // 'p' for portrait, 'mm' for millimeters, 'a4' for A4 size

        const imgDataPromises = currentCharts.map(chart => {
            return new Promise(resolve => {
                chart.canvas.toBlob(blob => {
                    const reader = new FileReader();
                    reader.onloadend = () => resolve(reader.result);
                    reader.readAsDataURL(blob);
                }, 'image/png');
            });
        });

        const imgDatas = await Promise.all(imgDataPromises);

        const imgWidth = 180; // mm
        const imgHeight = 100; // mm
        let yPos = 10;
        const margin = 10;

        // Add a title page or header for the PDF
        doc.setFont("helvetica", "bold");
        doc.setFontSize(22);
        // تنظیم RTL برای متن فارسی
        doc.setRTLTextPlugin("https://unpkg.com/jspdf-rtl-text@1.0.0/dist/jspdf.plugin.rtl-text.js", () => {
            doc.text("گزارش پروفیل بارهای مشترکین", doc.internal.pageSize.getWidth() / 2, yPos + 10, { align: 'center' });
            yPos += 30; // Move down for first chart

            for (let i = 0; i < imgDatas.length; i++) {
                const imgData = imgDatas[i];
                const customerName = parsedData[i] ? parsedData[i].name : 'نام مشترک نامشخص';
                const customerId = parsedData[i] ? parsedData[i].id : '';

                // Check if adding the next chart will exceed page height, if so, add new page
                if (yPos + imgHeight + margin + 15 > doc.internal.pageSize.getHeight()) { // +15 for title
                    doc.addPage();
                    yPos = margin;
                }

                doc.setFont("helvetica", "normal");
                doc.setFontSize(14);
                // تنظیم RTL برای متن فارسی
                doc.text(`نام مشترک: ${customerName} (شناسه: ${customerId})`, doc.internal.pageSize.getWidth() / 2, yPos + 5, { align: 'center' });
                yPos += 10; // Space for customer name

                doc.addImage(imgData, 'PNG', (doc.internal.pageSize.getWidth() - imgWidth) / 2, yPos, imgWidth, imgHeight);
                yPos += imgHeight + margin; // Move to the next position

                showProgress(10 + Math.floor((i + 1) / imgDatas.length * 80), `در حال آماده‌سازی PDF (${i + 1}/${imgDatas.length})...`);
            }

            doc.save("پروفیل_بارهای_مشترکین.pdf");
            hideProgress();
            Swal.fire('موفقیت', 'فایل PDF با موفقیت ایجاد شد.', 'success');
        });
    });

    // ================================================================================================
    // توابع کمکی DOM و منطق برنامه
    // ================================================================================================

    /**
     * برگه‌های موجود در فایل اکسل را به دراپ‌داون اضافه می‌کند.
     * @param {string[]} sheetNames - آرایه‌ای از نام برگه‌ها.
     */
    function updateSheetSelect(sheetNames) {
        if (sheetSelect) {
            sheetSelect.innerHTML = ''; // پاک کردن گزینه‌های قبلی
            sheetNames.forEach(sheetName => {
                const option = document.createElement('option');
                option.value = sheetName;
                option.textContent = sheetName;
                sheetSelect.appendChild(option);
            });
        }
    }

    /**
     * داده‌های اکسل را پردازش کرده و اطلاعات مورد نیاز (مصرف کل، پیک، زمان پیک، بار عصر، کاهش) را استخراج می‌کند.
     * این تابع با ساختار جدید فایل اکسل (15 دقیقه‌ای) سازگار شده است.
     * @param {Array<Array<any>>} jsonData - داده‌های خام JSON از اکسل.
     * @returns {Array<Object>} آرایه‌ای از آبجکت‌های پردازش شده مشترکین.
     */
    function processExcelData(jsonData) {
        const processed = [];
        if (!jsonData || jsonData.length < 2) return processed;

        const headerRow = jsonData[0].map(h => typeof h === 'string' ? h.trim() : ''); // اطمینان از string و trim کردن
        const dataRows = jsonData.slice(1);

        // شناسایی ستون‌های ID و نام مشترک با نام‌های احتمالی
        let idColIndex = headerRow.findIndex(h => h.toLowerCase() === 'id' || h.toLowerCase() === 'customer id' || h.toLowerCase() === 'customerid');
        let nameColIndex = headerRow.findIndex(h => h.toLowerCase() === 'نام مشترک' || h.toLowerCase() === 'customer name' || h.toLowerCase() === 'customername');
        // شناسایی ستون Contracted demand برای فیلتر حداکثر درصد کاهش
        let contractedDemandColIndex = headerRow.findIndex(h => h.toLowerCase() === 'contracted demand' || h.toLowerCase() === 'تقاضای قراردادی');

        if (idColIndex === -1 || nameColIndex === -1) {
            Swal.fire('خطا', 'ستون‌های "ID" یا "Customer id" و "نام مشترک" یا "Customer name" در فایل اکسل یافت نشدند.', 'error');
            return processed;
        }

        // شناسایی ستون‌های ساعتی 15 دقیقه‌ای
        const quarterHourColIndices = [];
        for (let i = 0; i < headerRow.length; i++) {
            const header = headerRow[i];
            // regex برای مطابقت با الگوهای "00:00 to 00:15 [KW]"
            if (header && header.match(/^\d{2}:\d{2} to \d{2}:\d{2} \[KW\]$/i)) {
                quarterHourColIndices.push(i);
            }
        }

        // انتظار 96 ستون (24 ساعت * 4 ربع ساعت)
        if (quarterHourColIndices.length < 96) {
            Swal.fire('خطا', `تعداد ستون‌های ساعتی 15 دقیقه‌ای ناکافی است. انتظار 96 ستون ساعتی (00:00 to 00:15 [KW] تا 23:45 to 00:00 [KW]) داریم، اما ${quarterHourColIndices.length} ستون یافت شد.`, 'error');
            return processed;
        }
        if (quarterHourColIndices.length > 96) {
             Swal.fire('اخطار', `تعداد ستون‌های ساعتی 15 دقیقه‌ای بیش از حد انتظار است. 96 ستون برای 24 ساعت کافی است. ممکن است برخی ستون‌ها نادیده گرفته شوند.`, 'warning');
        }

        const maxReductionEnabled = chkMaxReduction.checked;
        const maxReductionValue = parseFloat(txtMaxReduction.value);
        const minEveningLoadEnabled = chkMinEveningLoad.checked;
        const minEveningLoadValue = parseFloat(txtMinEveningLoad.value);

        dataRows.forEach(row => {
            const id = row[idColIndex];
            const name = row[nameColIndex];

            // اعتبارسنجی اولیه ID و نام
            if (id === null || id === undefined || String(id).trim() === '' ||
                name === null || name === undefined || String(name).trim() === '') {
                // اگر ID یا نام وجود نداشت یا خالی بود، این ردیف را نادیده بگیر
                return;
            }

            let totalConsumption = 0;
            let peakConsumption = 0;
            let peakTime = ''; // زمان پیک ساعتی
            let eveningLoad = 0; // بار عصر ساعتی
            let morningLoad = 0; // بار صبح ساعتی

            const hourlyLoads = new Array(24).fill(0); // برای نگهداری 24 مقدار ساعتی

            // پردازش داده‌های 15 دقیقه‌ای و تبدیل به ساعتی
            for (let h = 0; h < 24; h++) { // برای هر ساعت
                let sumOfQuarterLoads = 0;
                let quarterCount = 0;
                for (let q = 0; q < 4; q++) { // برای هر ربع ساعت در یک ساعت
                    const globalIndex = (h * 4) + q;
                    if (globalIndex < quarterHourColIndices.length) {
                        const colIdx = quarterHourColIndices[globalIndex];
                        let load = parseFloat(row[colIdx]);

                        // اگر بار نامعتبر بود، آن را 0 در نظر بگیرید تا محاسبات متوقف نشود
                        if (isNaN(load)) {
                            load = 0;
                        }
                        sumOfQuarterLoads += load;
                        quarterCount++;
                    }
                }
                if (quarterCount > 0) {
                    hourlyLoads[h] = sumOfQuarterLoads / quarterCount; // میانگین 4 ربع ساعت برای هر ساعت
                } else {
                    hourlyLoads[h] = 0; // اگر هیچ داده ربع ساعتی معتبری نبود، صفر بگذار
                }
                totalConsumption += hourlyLoads[h]; // جمع کل مصرف ساعتی (جمع KW ساعتی که معادل KWH است)
            }

            // پیدا کردن پیک و زمان پیک از داده‌های ساعتی
            if (hourlyLoads.length > 0) {
                peakConsumption = Math.max(...hourlyLoads);
                const peakHourIndex = hourlyLoads.indexOf(peakConsumption);
                peakTime = `${peakHourIndex.toString().padStart(2, '0')}:00`;
            }

            // محاسبه بار صبح بر اساس تنظیمات
            if (chkMorning.checked) {
                const morningHour = parseInt(txtMorning.value, 10);
                if (!isNaN(morningHour) && morningHour >= 0 && morningHour <= 23) {
                    morningLoad = hourlyLoads[morningHour] || 0;
                }
            }

            // محاسبه بار عصر بر اساس تنظیمات
            if (chkEvening.checked) {
                const eveningHour = parseInt(txtEvening.value, 10);
                if (!isNaN(eveningHour) && eveningHour >= 0 && eveningHour <= 23) {
                    eveningLoad = hourlyLoads[eveningHour] || 0;
                }
            }

            let reductionAmount = 0;
            let reductionPercent = 0;

            // محاسبه میزان کاهش و درصد کاهش (اگر بار صبح و عصر فعال باشد)
            if (chkMorning.checked && chkEvening.checked) {
                reductionAmount = morningLoad - eveningLoad;
                if (morningLoad > 0) {
                    reductionPercent = (reductionAmount / morningLoad) * 100;
                }
            }

            // --- منطق فیلتر: حداکثر درصد کاهش پیک ---
            if (maxReductionEnabled && !isNaN(maxReductionValue) && contractedDemandColIndex !== -1) {
                const contractedDemand = parseFloat(row[contractedDemandColIndex]);
                if (!isNaN(contractedDemand) && contractedDemand > 0) {
                    // محاسبه درصد کاهش واقعی نسبت به تقاضای قراردادی (پیک مصرف)
                    const actualReductionFromPeak = ((contractedDemand - peakConsumption) / contractedDemand) * 100;
                    if (actualReductionFromPeak > maxReductionValue) {
                        return; // این مشترک را حذف کن
                    }
                }
            }

            // --- منطق فیلتر: حذف اگر بار عصر کمتر از ۱۵KW ---
            if (minEveningLoadEnabled && !isNaN(minEveningLoadValue)) {
                if (eveningLoad < minEveningLoadValue) {
                    return; // این مشترک را حذف کن
                }
            }

            processed.push({
                id: String(id).trim(),
                name: String(name).trim(),
                totalConsumption: totalConsumption.toFixed(2), // 2 رقم اعشار
                peakConsumption: peakConsumption.toFixed(2),
                peakTime: peakTime,
                morningLoad: morningLoad.toFixed(2), // اضافه شدن بار صبح
                eveningLoad: eveningLoad.toFixed(2),
                reductionAmount: reductionAmount.toFixed(2),
                reductionPercent: reductionPercent.toFixed(2),
                hourlyLoads: hourlyLoads // برای نمودارها (داده‌های ساعتی)
            });
        });

        return processed;
    }

    /**
     * نتایج پردازش شده را در جدول نمایش می‌دهد.
     * @param {Array<Object>} data - آرایه‌ای از آبجکت‌های پردازش شده مشترکین.
     */
    function displayResultsTable(data) {
        if (!resultsTableBody) {
            console.error("resultsTableBody element not found.");
            return;
        }
        resultsTableBody.innerHTML = ''; // پاکسازی جدول قبلی
        if (data.length === 0) {
            // اطمینان از وجود resultsContainer قبل از دسترسی به style
            if (resultsContainer) resultsContainer.style.display = 'none';
            return;
        }

        data.forEach((item, index) => {
            const row = resultsTableBody.insertRow();
            row.insertCell().textContent = index + 1;
            row.insertCell().textContent = item.name;
            row.insertCell().textContent = item.id;
            row.insertCell().textContent = item.totalConsumption;
            row.insertCell().textContent = item.peakConsumption;
            row.insertCell().textContent = item.peakTime;
            row.insertCell().textContent = item.morningLoad; // نمایش بار صبح
            row.insertCell().textContent = item.eveningLoad;
            row.insertCell().textContent = item.reductionAmount;
            row.insertCell().textContent = item.reductionPercent;
        });
        // اطمینان از وجود resultsContainer قبل از دسترسی به style
        if (resultsContainer) resultsContainer.style.display = 'block'; // نمایش جدول
    }

    /**
     * نمودارهای پروفیل بار را برای هر مشترک رسم می‌کند.
     * @param {Array<Object>} data - آرایه‌ای از آبجکت‌های پردازش شده مشترکین.
     */
    function renderCharts(data) {
        if (!chartsContainer) {
            console.error("chartsContainer element not found.");
            return;
        }

        if (data.length === 0) {
            showNoChartsMessage();
            return;
        }

        hideNoChartsMessage();
        // ابتدا تمام فرزندان به جز h2 را حذف کنید تا اگر قبلاً پیامی بود، پاک شود.
        const chartsContainerChildren = Array.from(chartsContainer.children);
        chartsContainerChildren.forEach(child => {
            if (child.tagName !== 'H2' && child.id !== 'noChartsMessage') { // Keep the H2 title and noChartsMessage
                child.remove();
            }
        });


        data.forEach(customer => {
            const chartWrapper = document.createElement('div');
            chartWrapper.className = 'chart-wrapper';

            const chartTitle = document.createElement('h3');
            chartTitle.textContent = `${customer.name} (شناسه: ${customer.id})`;
            chartWrapper.appendChild(chartTitle);

            const canvas = document.createElement('canvas');
            canvas.id = `chart-${customer.id}`;
            chartWrapper.appendChild(canvas);

            const downloadBtn = document.createElement('button');
            downloadBtn.className = 'chart-download-btn';
            downloadBtn.textContent = 'دانلود نمودار';
            downloadBtn.addEventListener('click', () => {
                const img = canvas.toDataURL('image/png');
                const a = document.createElement('a');
                a.href = img;
                a.download = `پروفیل_بار_${customer.name}_${customer.id}.png`;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
            });
            chartWrapper.appendChild(downloadBtn);

            chartsContainer.appendChild(chartWrapper);

            // Chart.js Configuration
            const ctx = canvas.getContext('2d');
            // برچسب‌های ساعتی برای نمودار (00:00 تا 23:00)
            const labels = Array.from({ length: 24 }, (_, i) => `${i.toString().padStart(2, '0')}:00`);

            const annotations = {};

            // Add annotation for morning load if enabled
            if (chkMorning.checked && !isNaN(parseInt(txtMorning.value, 10))) {
                const morningHour = parseInt(txtMorning.value, 10);
                annotations.morningLoad = {
                    type: 'line',
                    borderColor: 'rgb(255, 99, 132)', // Changed color for morning
                    borderWidth: 2,
                    borderDash: [6, 6],
                    xMin: labels[morningHour],
                    xMax: labels[morningHour],
                    label: {
                        enabled: true,
                        content: `بار صبح (${customer.morningLoad} KW)`,
                        position: 'start',
                        backgroundColor: 'rgba(255, 99, 132, 0.8)',
                        font: { size: 10 }
                    }
                };
            }

            // Add annotation for evening load if enabled
            if (chkEvening.checked && !isNaN(parseInt(txtEvening.value, 10))) {
                const eveningHour = parseInt(txtEvening.value, 10);
                annotations.eveningLoad = {
                    type: 'line',
                    borderColor: 'rgb(54, 162, 235)', // Changed color for evening
                    borderWidth: 2,
                    borderDash: [6, 6],
                    xMin: labels[eveningHour],
                    xMax: labels[eveningHour],
                    label: {
                        enabled: true,
                        content: `بار عصر (${customer.eveningLoad} KW)`,
                        position: 'start',
                        backgroundColor: 'rgba(54, 162, 235, 0.8)',
                        font: { size: 10 }
                    }
                };
            }

            const chart = new Chart(ctx, {
                type: 'line',
                data: {
                    labels: labels, // استفاده از 24 برچسب ساعتی
                    datasets: [{
                        label: 'پروفیل بار (KW)',
                        data: customer.hourlyLoads, // داده‌های ساعتی
                        borderColor: 'rgba(75, 192, 192, 1)',
                        backgroundColor: 'rgba(75, 192, 192, 0.2)',
                        borderWidth: 2,
                        fill: true,
                        tension: 0.3
                    },
                    {
                        label: `پیک مصرف (${customer.peakConsumption} KW)`,
                        data: Array(24).fill(customer.peakConsumption),
                        borderColor: 'rgba(255, 159, 64, 1)',
                        borderWidth: 2,
                        borderDash: [5, 5],
                        pointRadius: 0,
                        fill: false
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        title: {
                            display: false, // Title is in h3 element
                        },
                        legend: {
                            display: true,
                            position: 'bottom',
                            labels: {
                                usePointStyle: true,
                                generateLabels: function(chart) {
                                    const datasets = chart.data.datasets;
                                    const customLabels = datasets.map((dataset, i) => ({
                                        text: dataset.label,
                                        fillStyle: dataset.backgroundColor || dataset.borderColor,
                                        strokeStyle: dataset.borderColor,
                                        lineWidth: dataset.borderWidth,
                                        hidden: !chart.isDatasetVisible(i),
                                        index: i
                                    }));

                                    // Add custom labels for annotations if enabled
                                    if (annotations.morningLoad && annotations.morningLoad.label.enabled) {
                                        customLabels.push({
                                            text: `بار صبح (${customer.morningLoad} KW)`,
                                            fillStyle: annotations.morningLoad.borderColor,
                                            strokeStyle: annotations.morningLoad.borderColor,
                                            lineWidth: annotations.morningLoad.borderWidth,
                                            lineDash: annotations.morningLoad.borderDash,
                                            hidden: false,
                                            index: customLabels.length // unique index
                                        });
                                    }
                                    if (annotations.eveningLoad && annotations.eveningLoad.label.enabled) {
                                        customLabels.push({
                                            text: `بار عصر (${customer.eveningLoad} KW)`,
                                            fillStyle: annotations.eveningLoad.borderColor,
                                            strokeStyle: annotations.eveningLoad.borderColor,
                                            lineWidth: annotations.eveningLoad.borderWidth,
                                            lineDash: annotations.eveningLoad.borderDash,
                                            hidden: false,
                                            index: customLabels.length // unique index
                                        });
                                    }
                                    return customLabels;
                                }
                            }
                        },
                        tooltip: {
                            mode: 'index',
                            intersect: false,
                            callbacks: {
                                title: function(context) {
                                    // Use the hour from the label directly
                                    return `ساعت: ${context[0].label}`;
                                },
                                label: function(context) {
                                    let label = context.dataset.label || '';
                                    if (label) {
                                        label += ': ';
                                    }
                                    if (context.parsed.y !== null) {
                                        label += parseFloat(context.parsed.y).toFixed(2) + ' KW';
                                    }
                                    return label;
                                }
                            },
                            rtl: true // Enable RTL for tooltips
                        },
                        annotation: {
                            annotations: annotations // Directly use the annotations object
                        }
                    },
                    scales: {
                        x: {
                            title: {
                                display: true,
                                text: 'ساعت شبانه‌روز'
                            },
                            grid: {
                                display: false
                            }
                        },
                        y: {
                            title: {
                                display: true,
                                text: 'مصرف (KW)'
                            },
                            beginAtZero: true
                        }
                    }
                }
            });
            currentCharts.push(chart); // ذخیره نمونه نمودار برای پاکسازی بعدی
        });
    }

});