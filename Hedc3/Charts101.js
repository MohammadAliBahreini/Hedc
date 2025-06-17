// ====================================================================================================
// تعریف متغیرهای سراسری
// ====================================================================================================
let workbook;
let parsedData = [];
let currentCharts = [];
let filteredData = []; // برای نگهداری داده‌های فیلتر شده
let appLogs = []; // آرایه برای نگهداری لاگ‌های برنامه

// ====================================================================================================
// توابع کمکی
// ====================================================================================================

/**
 * تابع لاگ‌گیری حرفه‌ای
 * @param {string} level - سطح لاگ (info, warn, error, debug)
 * @param {string} message - پیام لاگ
 * @param {object} context - اطلاعات اضافی مربوط به لاگ (اختیاری)
 */
function log(level, message, context = {}) {
    const timestamp = new Date().toISOString();
    const logEntry = {
        timestamp: timestamp,
        level: level.toUpperCase(),
        message: message,
        context: context
    };
    appLogs.push(logEntry); // ذخیره لاگ در آرایه
    console.log(`[${timestamp}] [${level.toUpperCase()}]: ${message}`, context); // نمایش در کنسول
}

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
        if (percent === 100) {
            // After a short delay, hide the progress container
            setTimeout(() => {
                progressContainer.style.opacity = '0';
                setTimeout(() => progressContainer.style.display = 'none', 500);
            }, 1000); // 1 second delay before starting fade out
        } else {
            progressContainer.style.display = 'block'; // Ensure it's visible when processing
            progressContainer.style.opacity = '1';
        }
    }
}

/**
 * تابع تبدیل زمان اکسل به تاریخ جاوااسکریپت
 * @param {number} excelTimestamp - زمان در فرمت اکسل
 * @returns {string} تاریخ فرمت شده (YYYY/MM/DD)
 */
function excelDateToJSDate(excelTimestamp) {
    if (typeof excelTimestamp !== 'number' || isNaN(excelTimestamp)) {
        return ''; // بازگرداندن رشته خالی برای مقادیر نامعتبر
    }
    const date = new Date(Math.round((excelTimestamp - 25569) * 86400 * 1000));
    const year = date.getFullYear();
    const month = (date.getMonth() + 1).toString().padStart(2, '0');
    const day = date.getDate().toString().padStart(2, '0');
    return `${year}/${month}/${day}`;
}


/**
 * تابع پاکسازی و نابود کردن نمودارهای قبلی
 */
function destroyCharts() {
    currentCharts.forEach(chart => {
        chart.destroy();
    });
    currentCharts = [];
    const chartsGrid = document.querySelector('.charts-grid');
    if (chartsGrid) {
        chartsGrid.innerHTML = ''; // پاک کردن همه div های نمودار
    }
    const noChartsMessage = document.getElementById('noChartsMessage');
    if (noChartsMessage) {
        noChartsMessage.style.display = 'block';
    }
    log('info', 'All existing charts destroyed.');
}

/**
 * تابع محاسبه بار صبح و عصر
 * این تابع بر روی 96 نقطه 15 دقیقه‌ای اصلی کار می‌کند تا دقت محاسبات آماری حفظ شود.
 * @param {Array<number>} hourlyData - آرایه داده‌های ساعتی مصرف (96 نقطه)
 * @param {string} startTime - زمان شروع (مثال: "06:00")
 * @param {string} endTime - زمان پایان (مثال: "09:00")
 * @param {string} calcType - نوع محاسبه ("avg", "max", "min")
 * @returns {number} مقدار محاسبه شده
 */
function calculatePeakLoad(hourlyData, startTime, endTime, calcType) {
    // Convert HH:MM to 15-minute intervals index (0-95)
    const getIndex = (timeString) => {
        const [hour, minute] = timeString.split(':').map(Number);
        return (hour * 4) + (minute / 15);
    };

    const startIndex = getIndex(startTime);
    const endIndex = getIndex(endTime);

    let relevantData = [];
    for (let i = startIndex; i <= endIndex; i++) {
        if (hourlyData[i] !== undefined && typeof hourlyData[i] === 'number' && !isNaN(hourlyData[i])) {
            relevantData.push(hourlyData[i]);
        }
    }

    if (relevantData.length === 0) {
        return 0;
    }

    switch (calcType) {
        case 'avg':
            return relevantData.reduce((sum, val) => sum + val, 0) / relevantData.length;
        case 'max':
            return Math.max(...relevantData);
        case 'min':
            return Math.min(...relevantData);
        default:
            return 0;
    }
}

/**
 * تابع رندر کردن نمودار برای یک مشترک خاص
 * @param {Object} customerData - داده‌های مصرفی مشترک
 * @param {boolean} isBatchRender - آیا این نمودار در یک فرایند دسته‌ای رسم می‌شود (برای مدیریت پاکسازی نمودارها)
 */
function renderChart(customerData, isBatchRender = false) {
    const chartsGrid = document.querySelector('.charts-grid');
    if (!chartsGrid) return;

    if (!isBatchRender) { // Only destroy charts if it's a single chart request
        destroyCharts();
    }

    const chartContainerDiv = document.createElement('div');
    chartContainerDiv.classList.add('chart-container');
    const canvas = document.createElement('canvas');
    chartContainerDiv.appendChild(canvas);
    chartsGrid.appendChild(chartContainerDiv);

    const noChartsMessage = document.getElementById('noChartsMessage');
    if (noChartsMessage) {
        noChartsMessage.style.display = 'none';
    }

    // Generate labels for 24 hours (H01 to H24)
    const hourlyLabels = Array.from({ length: 24 }, (_, i) => `H${String(i + 1).padStart(2, '0')}`);
    
    // Aggregate 96 15-minute data points into 24 hourly points (average)
    const aggregatedHourlyData = [];
    for (let h = 0; h < 24; h++) {
        let sum = 0;
        let count = 0;
        for (let i = 0; i < 4; i++) { // Each hour has 4 15-minute points
            const dataIndex = (h * 4) + i;
            if (customerData.hourlyData[dataIndex] !== undefined && typeof customerData.hourlyData[dataIndex] === 'number' && !isNaN(customerData.hourlyData[dataIndex])) {
                sum += customerData.hourlyData[dataIndex];
                count++;
            }
        }
        aggregatedHourlyData.push(count > 0 ? sum / count : 0);
    }

    // Peak calculation for annotations will now use the 24-point aggregated data for consistency with chart labels
    // Define peak times in terms of hours for this new chart structure (1-24)
    const morningPeakStartHour = parseInt(document.getElementById('morningPeakStart').value.split(':')[0]);
    const morningPeakEndHour = parseInt(document.getElementById('morningPeakEnd').value.split(':')[0]);
    const eveningPeakStartHour = parseInt(document.getElementById('eveningPeakStart').value.split(':')[0]);
    const eveningPeakEndHour = parseInt(document.getElementById('eveningPeakEnd').value.split(':')[0]);
    const morningCalcType = document.getElementById('morningCalcType').value;
    const eveningCalcType = document.getElementById('eveningCalcType').value;

    const getHourlyPeakLoad = (hourlyDataArray, startHour, endHour, calcType) => {
        let relevantData = [];
        // Convert 1-indexed hours (H1-H24) to 0-indexed array indices (0-23)
        const startIndex = startHour - 1;
        const endIndex = endHour - 1;

        for (let i = startIndex; i <= endIndex; i++) {
            if (hourlyDataArray[i] !== undefined && typeof hourlyDataArray[i] === 'number' && !isNaN(hourlyDataArray[i])) {
                relevantData.push(hourlyDataArray[i]);
            }
        }
        if (relevantData.length === 0) return 0;

        switch (calcType) {
            case 'avg': return relevantData.reduce((sum, val) => sum + val, 0) / relevantData.length;
            case 'max': return Math.max(...relevantData);
            case 'min': return Math.min(...relevantData);
            default: return 0;
        }
    };

    const morningPeakValue = getHourlyPeakLoad(aggregatedHourlyData, morningPeakStartHour, morningPeakEndHour, morningCalcType);
    const eveningPeakValue = getHourlyPeakLoad(aggregatedHourlyData, eveningPeakStartHour, eveningPeakEndHour, eveningCalcType);

    // Find the exact time labels for peak values in the H-labels
    const getHourlyPeakTimeLabel = (value, dataArray, labelsArray, startHour, endHour) => {
        const startIndex = startHour - 1;
        const endIndex = endHour - 1;
        for (let i = startIndex; i <= endIndex; i++) {
            if (dataArray[i] === value) {
                return labelsArray[i];
            }
        }
        return '';
    };

    const morningPeakTimeLabel = getHourlyPeakTimeLabel(morningPeakValue, aggregatedHourlyData, hourlyLabels, morningPeakStartHour, morningPeakEndHour);
    const eveningPeakTimeLabel = getHourlyPeakTimeLabel(eveningPeakValue, aggregatedHourlyData, hourlyLabels, eveningPeakStartHour, eveningPeakEndHour);

    const annotations = {};
    if (morningPeakValue > 0) {
        annotations.morningPeak = {
            type: 'point',
            xValue: morningPeakTimeLabel || `H${String(morningPeakStartHour).padStart(2, '0')}`,
            yValue: morningPeakValue,
            backgroundColor: 'rgba(255, 99, 132, 0.8)',
            radius: 5,
            pointStyle: 'circle',
            label: {
                content: `بار صبح: ${morningPeakValue.toFixed(0)}kW (${morningPeakTimeLabel || `H${String(morningPeakStartHour).padStart(2, '0')}`})`,
                display: true,
                position: 'top',
                backgroundColor: 'rgba(255, 99, 132, 0.7)',
                font: {
                    size: 10,
                    weight: 'bold',
                    family: 'Vazirmatn'
                },
                color: '#fff'
            }
        };
    }
    if (eveningPeakValue > 0) {
        annotations.eveningPeak = {
            type: 'point',
            xValue: eveningPeakTimeLabel || `H${String(eveningPeakStartHour).padStart(2, '0')}`,
            yValue: eveningPeakValue,
            backgroundColor: 'rgba(54, 162, 235, 0.8)',
            radius: 5,
            pointStyle: 'circle',
            label: {
                content: `بار عصر: ${eveningPeakValue.toFixed(0)}kW (${eveningPeakTimeLabel || `H${String(eveningPeakStartHour).padStart(2, '0')}`})`,
                display: true,
                position: 'top',
                backgroundColor: 'rgba(54, 162, 235, 0.7)',
                font: {
                    size: 10,
                    weight: 'bold',
                    family: 'Vazirmatn'
                },
                color: '#fff'
            }
        };
    }

    const chart = new Chart(canvas, {
        type: 'line',
        data: {
            labels: hourlyLabels, // Use H01-H24 labels
            datasets: [{
                label: `پروفیل بار مشترک: ${customerData.customerName} ( ${customerData.customerId} - ${customerData.serialNo})`,
                data: aggregatedHourlyData, // Use aggregated data for the chart
                borderColor: 'rgba(75, 192, 192, 1)',
                backgroundColor: 'rgba(75, 192, 192, 0.2)',
                borderWidth: 1,
                fill: true,
                tension: 0.4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: true,
                    position: 'top',
                    labels: {
                        font: {
                            family: 'Vazirmatn',
                        }
                    }
                },
                title: {
                    display: true,
                    text: `پروفیل بار روزانه برای ${customerData.customerName}`,
                    font: {
                        family: 'Vazirmatn',
                        size: 14
                    }
                },
                tooltip: {
                    rtl: true,
                    bodyFont: {
                        family: 'Vazirmatn'
                    },
                    titleFont: {
                        family: 'Vazirmatn'
                    },
                    callbacks: {
                        title: function(context) {
                            return `ساعت: ${context[0].label.replace('H', '')}`; // Show just the hour number
                        },
                        label: function(context) {
                            return `مصرف: ${context.parsed.y} kW`;
                        }
                    }
                },
                annotation: {
                    annotations: annotations
                }
            },
            scales: {
                x: {
                    title: {
                        display: true,
                        text: 'ساعت',
                        font: {
                            family: 'Vazirmatn'
                        }
                    },
                    ticks: {
                        maxRotation: 0, // Keep labels horizontal
                        minRotation: 0,
                        font: {
                            family: 'Vazirmatn'
                        },
                        autoSkipPadding: 0,
                        // Show all 24 labels
                        callback: function(val, index) {
                            return hourlyLabels[index]; // Display H01, H02, ...
                        }
                    }
                },
                y: {
                    title: {
                        display: true,
                        text: 'مصرف (KW)',
                        font: {
                            family: 'Vazirmatn'
                        }
                    },
                    beginAtZero: true,
                    ticks: {
                        font: {
                            family: 'Vazirmatn'
                        }
                    }
                }
            }
        }
    });
    currentCharts.push(chart);
    log('info', `Chart rendered for customer: ${customerData.customerName}`);
}

/**
 * تابع رندر کردن همه نمودارها به صورت دسته‌ای
 */
async function renderAllCharts() {
    if (filteredData.length === 0) {
        Swal.fire('توجه', 'داده‌ای برای رسم نمودارها وجود ندارد. ابتدا داده‌ها را پردازش کنید.', 'warning');
        log('warn', 'Attempted to render all charts with no data.');
        return;
    }

    if (filteredData.length > 50) { // Limit for performance warning
        const result = await Swal.fire({
            title: 'هشدار عملکرد!',
            text: `تعداد مشترکین (${filteredData.length}) زیاد است. رسم همه نمودارها ممکن است باعث کندی شدید یا فریز شدن مرورگر شود. آیا مطمئنید؟`,
            icon: 'warning',
            showCancelButton: true,
            confirmButtonColor: '#3085d6',
            cancelButtonColor: '#d33',
            confirmButtonText: 'بله، ادامه بده!',
            cancelButtonText: 'خیر'
        });
        if (!result.isConfirmed) {
            log('info', 'Rendering all charts cancelled by user due to high volume.');
            return;
        }
    }

    destroyCharts(); // Clear existing charts before drawing all
    log('info', `Attempting to render all ${filteredData.length} charts.`);

    showProgress(0, 'در حال رسم همه نمودارها...');
    const chartsContainer = document.getElementById('chartsContainer');
    if (chartsContainer) {
        const noChartsMessage = document.getElementById('noChartsMessage');
        if (noChartsMessage) noChartsMessage.style.display = 'none';
    }

    // Process charts with a small delay to avoid freezing the UI
    for (let i = 0; i < filteredData.length; i++) {
        renderChart(filteredData[i], true); // Pass true for isBatchRender
        showProgress(Math.round(((i + 1) / filteredData.length) * 100), `در حال رسم نمودار ${i + 1} از ${filteredData.length}`);
        await new Promise(resolve => setTimeout(resolve, 10)); // Small delay for UI to breathe
    }
    showProgress(100, 'رسم همه نمودارها کامل شد.');
    log('info', `Finished rendering all ${filteredData.length} charts.`);
}


/**
 * تابع نمایش داده‌های پردازش شده در جدول HTML
 * @param {Array} dataToDisplay - آرایه‌ای از داده‌ها برای نمایش
 */
function displayParsedData(dataToDisplay) {
    const resultsTableBody = document.querySelector('# tbody');
    if (!resultsTableBody) {
        log('error', 'resultsTableBody element not found. Cannot display data.');
        return;
    }
    resultsTableBody.innerHTML = ''; // پاک کردن محتوای قبلی

    if (dataToDisplay.length === 0) {
        const noDataRow = resultsTableBody.insertRow();
        const cell = noDataRow.insertCell();
        cell.colSpan = 14; // تعداد ستون‌ها
        cell.textContent = 'داده‌ای برای نمایش وجود ندارد.';
        cell.style.textAlign = 'center';
        log('info', 'No data to display in the table.');
        return;
    }

    dataToDisplay.forEach((rowData, index) => {
        const row = resultsTableBody.insertRow();
        row.insertCell().textContent = index + 1; // شماره ردیف (از 1)
        row.insertCell().textContent = rowData.serialNo;
        row.insertCell().textContent = rowData.date; // اضافه شدن تاریخ
        row.insertCell().textContent = rowData.customerName;
        row.insertCell().textContent = rowData.billingId;
        row.insertCell().textContent = rowData.customerId;
        row.insertCell().textContent = rowData.contractedDemand;
        row.insertCell().textContent = rowData.address;

        // Calculate and display Morning and Evening Peak Loads
        // Use the original 15-min interval peak calculation as it's more precise for table data
        const morningPeakStart = document.getElementById('morningPeakStart').value;
        const morningPeakEnd = document.getElementById('morningPeakEnd').value;
        const eveningPeakStart = document.getElementById('eveningPeakStart').value;
        const eveningPeakEnd = document.getElementById('eveningPeakEnd').value;
        const morningCalcType = document.getElementById('morningCalcType').value;
        const eveningCalcType = document.getElementById('eveningCalcType').value;

        const morningLoad = calculatePeakLoad(rowData.hourlyData, morningPeakStart, morningPeakEnd, morningCalcType).toFixed(2);
        const eveningLoad = calculatePeakLoad(rowData.hourlyData, eveningPeakStart, eveningPeakEnd, eveningCalcType).toFixed(2);

        row.insertCell().textContent = morningLoad;
        row.insertCell().textContent = eveningLoad;

        // Calculate Reduction
        const contractedDemand = parseFloat(rowData.contractedDemand);
        let reductionAmount = 0;
        let reductionPercentage = 0;

        if (contractedDemand > 0) {
            // Ensure hourlyData contains valid numbers before finding max
            const validHourlyData = rowData.hourlyData.filter(val => typeof val === 'number' && !isNaN(val));
            const maxLoad = validHourlyData.length > 0 ? Math.max(...validHourlyData) : 0;
            
            reductionAmount = contractedDemand - maxLoad;
            reductionPercentage = (reductionAmount / contractedDemand) * 100;
        }

        row.insertCell().textContent = reductionAmount.toFixed(2);
        row.insertCell().textContent = reductionPercentage.toFixed(2) + '%';


        const chartActionsCell = row.insertCell(); // Cell for chart button
        const chartBtn = document.createElement('button');
        chartBtn.classList.add('btn-chart');
        chartBtn.innerHTML = '<i class="fas fa-chart-line"></i>';
        chartBtn.title = 'نمایش نمودار تکی';
        chartBtn.onclick = () => {
            // This will call renderChart for a single customer and clear others
            renderChart(rowData, false); // Pass false for single chart request
            log('info', `Chart button clicked for customer: ${rowData.customerName}`);
        };
        chartActionsCell.appendChild(chartBtn);

        const deleteActionsCell = row.insertCell(); // Cell for delete button
        const deleteBtn = document.createElement('button');
        deleteBtn.classList.add('btn-delete');
        deleteBtn.innerHTML = '<i class="fas fa-trash"></i>';
        deleteBtn.title = 'حذف رکورد';
        deleteBtn.onclick = () => handleDeleteRow(rowData.originalIndex); // استفاده از originalIndex برای حذف
        deleteActionsCell.appendChild(deleteBtn);
    });
    log('info', `Displayed ${dataToDisplay.length} rows in the table.`);
}

/**
 * تابع حذف یک ردیف از parsedData و به‌روزرسانی جدول
 * @param {number} originalIndex - ایندکس اصلی رکورد در parsedData
 */
function handleDeleteRow(originalIndex) {
    Swal.fire({
        title: 'آیا مطمئن هستید؟',
        text: "این عملیات قابل بازگشت نیست!",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#3085d6',
        cancelButtonColor: '#d33',
        confirmButtonText: 'بله، حذف کن!',
        cancelButtonText: 'خیر'
    }).then((result) => {
        if (result.isConfirmed) {
            // Remove from parsedData (the original source)
            parsedData = parsedData.filter(item => item.originalIndex !== originalIndex);

            // Re-apply filter to update filteredData
            applyFilter(); // This will also call displayParsedData and re-number rows

            destroyCharts(); // Clear charts after deletion
            Swal.fire(
                'حذف شد!',
                'رکورد با موفقیت حذف شد.',
                'success'
            );
            log('info', `Row with original index ${originalIndex} deleted.`);
        }
    }).catch(error => {
        log('error', 'Error confirming row deletion.', { error: error });
    });
}

/**
 * تابع پردازش فایل اکسل
 */
async function processExcelFile() {
    const fileInput = document.getElementById('excelFile');
    const sheetSelect = document.getElementById('sheetSelect');
    const processStatusMessage = document.getElementById('processStatusMessage');
    const processDataBtn = document.getElementById('processDataBtn');

    if (!fileInput || !fileInput.files || fileInput.files.length === 0) {
        Swal.fire('خطا', 'لطفا یک فایل اکسل انتخاب کنید.', 'error');
        log('error', 'No Excel file selected for processing.');
        return;
    }

    processDataBtn.disabled = true; // Disable button during processing
    showProgress(0, 'در حال خواندن فایل...');
    processStatusMessage.textContent = 'در حال خواندن فایل...';
    log('info', `Starting Excel file processing for: ${fileInput.files[0].name}`);

    const file = fileInput.files[0];
    const reader = new FileReader();

    reader.onload = async (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array' });
            log('info', 'Excel workbook loaded successfully.');

            // Populate sheet selection dropdown
            if (sheetSelect) {
                sheetSelect.innerHTML = '';
                workbook.SheetNames.forEach(sheetName => {
                    const option = document.createElement('option');
                    option.value = sheetName;
                    option.textContent = sheetName;
                    sheetSelect.appendChild(option);
                });
                sheetSelect.disabled = false;
                // Select the first sheet by default after populating
                if (workbook.SheetNames.length > 0) {
                    sheetSelect.value = workbook.SheetNames[0];
                }
            } else {
                log('error', 'Sheet select element not found.');
                Swal.fire('خطا', 'عنصر انتخاب شیت یافت نشد. لطفاً صفحه را رفرش کنید.', 'error');
                processStatusMessage.textContent = 'خطا در بارگذاری برنامه.';
                showProgress(0, 'خطا');
                return;
            }

            // After populating sheetSelect, now get the sheetName
            const sheetName = sheetSelect.value; // Now sheetSelect.value should be available
            if (!sheetName) {
                Swal.fire('خطا', 'هیچ شیتی در فایل اکسل یافت نشد.', 'error');
                processStatusMessage.textContent = 'خطا: فایل اکسل فاقد شیت است.';
                showProgress(0, 'خطا');
                log('error', 'No sheets found in the Excel workbook.');
                return;
            }
            
            processCurrentSheet(sheetName);

        } catch (error) {
            Swal.fire('خطا', `خطا در پردازش فایل: ${error.message}. لطفاً از صحت فرمت فایل و ساختار ستون‌ها اطمینان حاصل کنید.`, 'error');
            processStatusMessage.textContent = 'خطا در پردازش فایل.';
            showProgress(0, 'خطا در پردازش');
            log('error', 'Error processing Excel file.', { errorMessage: error.message, stack: error.stack });
        } finally {
            processDataBtn.disabled = false; // Re-enable button after processing
        }
    };

    reader.onerror = (error) => {
        Swal.fire('خطا', `خطا در خواندن فایل: ${error.message}`, 'error');
        processStatusMessage.textContent = 'خطا در خواندن فایل.';
        showProgress(0, 'خطا در خواندن');
        log('error', 'Error reading file.', { errorMessage: error.message, stack: error.stack });
        processDataBtn.disabled = false; // Re-enable button
    };

    reader.readAsArrayBuffer(file);
}

/**
 * تابع کمکی برای پردازش شیت فعلی
 * این تابع هم برای پردازش اولیه و هم برای تغییر شیت استفاده می‌شود.
 * @param {string} sheetName - نام شیتی که باید پردازش شود.
 */
async function processCurrentSheet(sheetName) { // Made async to allow await
    const processStatusMessage = document.getElementById('processStatusMessage');
    parsedData = []; // Clear previous data
    destroyCharts(); // Clear any existing charts

    if (!workbook || !workbook.Sheets[sheetName]) {
        Swal.fire('خطا', 'شیف انتخاب شده یافت نشد یا ورک‌بوک خالی است.', 'error');
        processStatusMessage.textContent = 'خطا: شیت نامعتبر.';
        showProgress(0, 'خطا');
        log('error', `Attempted to process non-existent sheet: ${sheetName}`);
        return;
    }

    try {
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        log('info', `Processing sheet '${sheetName}'. Rows found: ${json.length}`);

        if (json.length <= 1) { // If only header or empty
            Swal.fire('خطا', `شیف "${sheetName}" خالی است یا فرمت آن صحیح نیست (حداقل یک ردیف سربرگ و یک ردیف داده نیاز است).`, 'error');
            processStatusMessage.textContent = `خطا: شیف "${sheetName}" داده معتبر ندارد.`;
            showProgress(0, 'خطا در پردازش');
            log('error', `Sheet '${sheetName}' is empty or has invalid format (less than 2 rows).`, { jsonContent: json });
            displayParsedData([]); // Clear table
            return;
        }

        const dataRows = json.slice(1); // Assume first row is header
        let validRowsCount = 0;
        const expectedMinColumns = 8 + 96; // 8 for customer info + 96 for hourly data

        dataRows.forEach((row, originalIndex) => {
            if (row.length < expectedMinColumns) {
                log('warn', `Row ${originalIndex + 2} in sheet '${sheetName}' has insufficient columns (${row.length} < ${expectedMinColumns}), skipping.`, { rowData: row });
                return; // Skip this row if not enough columns
            }

            const customerInfo = {
                originalIndex: originalIndex, // To uniquely identify original record for deletion
                serialNo: row[1] !== undefined ? row[1] : '', // Column B (index 1)
                date: row[2] !== undefined ? excelDateToJSDate(row[2]) : '', // Column C (index 2)
                customerName: row[3] !== undefined ? row[3] : '', // Column D (index 3)
                billingId: row[4] !== undefined ? row[4] : '', // Column E (index 4)
                customerId: row[5] !== undefined ? row[5] : '', // Column F (index 5)
                contractedDemand: row[6] !== undefined ? parseFloat(row[6]) : 0, // Column G (index 6)
                address: row[7] !== undefined ? row[7] : '', // Column H (index 7)
            };

            const hourlyData = [];
            let allHourlyDataValid = true;
            for (let i = 8; i < expectedMinColumns; i++) {
                const value = parseFloat(row[i]);
                if (isNaN(value)) {
                    allHourlyDataValid = false; // Mark if any hourly data is not a number
                    hourlyData.push(0); // Push 0 or handle as needed, but mark as invalid
                } else {
                    hourlyData.push(value);
                }
            }
            
            // Validate hourly data length and if all parts are numbers
            if (hourlyData.length === 96 && allHourlyDataValid) {
                parsedData.push({ ...customerInfo, hourlyData: hourlyData });
                validRowsCount++;
            } else {
                log('warn', `Row ${originalIndex + 2} in sheet '${sheetName}' has invalid hourly data (length: ${hourlyData.length}, allValid: ${allHourlyDataValid}), skipping.`, { hourlyData: hourlyData, rowData: row });
            }
        });

        if (validRowsCount > 0) {
            filteredData = [...parsedData]; // Initialize filteredData with all parsed data
            displayParsedData(filteredData);

            // Enable export buttons and "render all charts" button
            document.getElementById('exportExcelBtn').disabled = false;
            document.getElementById('exportPdfBtn').disabled = false;
            document.getElementById('exportChartsAsImagesBtn').disabled = false;
            document.getElementById('renderAllChartsBtn').disabled = false; // Enable new button

            processStatusMessage.textContent = 'پردازش کامل شد.';
            showProgress(100, 'پردازش کامل شد.');
            Swal.fire('موفقیت!', `فایل اکسل با موفقیت پردازش شد. (${validRowsCount} رکورد معتبر یافت شد.)`, 'success');
            log('info', `Successfully processed ${validRowsCount} valid rows from sheet '${sheetName}'.`);

            // Automatically render all charts after successful processing
            await renderAllCharts(); // Use await here
            log('info', 'Automatically rendered all charts after sheet processing.');

        } else {
            Swal.fire('خطا', `داده معتبری با فرمت صحیح در شیف "${sheetName}" یافت نشد. لطفاً ساختار ستون‌ها را بررسی کنید (حداقل 104 ستون شامل اطلاعات مشترک و 96 نقطه ساعتی و مقادیر عددی صحیح).`, 'error');
            processStatusMessage.textContent = 'خطا: داده معتبر یافت نشد.';
            showProgress(0, 'خطا در پردازش');
            log('error', `No valid data found in Excel sheet '${sheetName}' after processing. Check column structure and data types.`);
            displayParsedData([]); // Clear table
            destroyCharts(); // Clear charts
        }
    } catch (error) {
        Swal.fire('خطا', `خطا در پردازش شیف "${sheetName}": ${error.message}. لطفاً از صحت فرمت فایل و ساختار ستون‌ها اطمینان حاصل کنید.`, 'error');
        processStatusMessage.textContent = `خطا در پردازش شیف "${sheetName}".`;
        showProgress(0, 'خطا در پردازش');
        log('error', `Error processing sheet '${sheetName}'.`, { errorMessage: error.message, stack: error.stack });
        displayParsedData([]); // Clear table
        destroyCharts(); // Clear charts
    }
}


/**
 * تابع اعمال فیلتر بر روی داده‌ها
 */
function applyFilter() {
    const filterColumn = document.getElementById('filterColumn').value;
    const filterValue = document.getElementById('filterValue').value.toLowerCase().trim();

    if (!filterValue) {
        filteredData = [...parsedData]; // If filter is empty, display all original parsed data
        log('info', 'Filter cleared, displaying all parsed data.');
    } else {
        filteredData = parsedData.filter(item => {
            let valueToMatch = '';
            switch (filterColumn) {
                case 'all':
                    // Search across multiple relevant text/number columns
                    return (item.customerName.toLowerCase().includes(filterValue) ||
                            item.billingId.toString().toLowerCase().includes(filterValue) ||
                            item.customerId.toString().toLowerCase().includes(filterValue) ||
                            item.address.toLowerCase().includes(filterValue) ||
                            item.serialNo.toString().toLowerCase().includes(filterValue) ||
                            item.date.toLowerCase().includes(filterValue) // Include date in 'all' search
                           );
                case 'serialNo':
                    valueToMatch = item.serialNo.toString();
                    break;
                case 'customerName':
                    valueToMatch = item.customerName;
                    break;
                case 'billingId':
                    valueToMatch = item.billingId.toString();
                    break;
                case 'customerId':
                    valueToMatch = item.customerId.toString();
                    break;
                case 'address':
                    valueToMatch = item.address;
                    break;
                default:
                    return false;
            }
            return valueToMatch.toLowerCase().includes(filterValue);
        });
        log('info', `Filter applied: Column='${filterColumn}', Value='${filterValue}'. Found ${filteredData.length} records.`);
    }
    displayParsedData(filteredData);
    destroyCharts(); // Clear charts after filter is applied/cleared
}

/**
 * تابع پاک کردن فیلتر
 */
function clearFilter() {
    document.getElementById('filterValue').value = '';
    document.getElementById('filterColumn').value = 'all';
    applyFilter(); // Apply filter without value to show all data
    log('info', 'Filter form cleared.');
}

/**
 * تابع خروجی اکسل
 */
function exportToExcel() {
    if (filteredData.length === 0) {
        Swal.fire('توجه', 'داده‌ای برای خروجی گرفتن وجود ندارد.', 'warning');
        log('warn', 'Attempted to export to Excel with no data.');
        return;
    }

    const exportData = filteredData.map(row => {
        const morningPeakStart = document.getElementById('morningPeakStart').value;
        const morningPeakEnd = document.getElementById('morningPeakEnd').value;
        const eveningPeakStart = document.getElementById('eveningPeakStart').value;
        const eveningPeakEnd = document.getElementById('eveningPeakEnd').value;
        const morningCalcType = document.getElementById('morningCalcType').value;
        const eveningCalcType = document.getElementById('eveningCalcType').value;

        const morningLoad = calculatePeakLoad(row.hourlyData, morningPeakStart, morningPeakEnd, morningCalcType).toFixed(2);
        const eveningLoad = calculatePeakLoad(row.hourlyData, eveningPeakStart, eveningPeakEnd, eveningCalcType).toFixed(2);

        const contractedDemand = parseFloat(row.contractedDemand);
        let reductionAmount = 0;
        let reductionPercentage = 0;
        if (contractedDemand > 0) {
            const validHourlyData = row.hourlyData.filter(val => typeof val === 'number' && !isNaN(val));
            const maxLoad = validHourlyData.length > 0 ? Math.max(...validHourlyData) : 0;
            reductionAmount = contractedDemand - maxLoad;
            reductionPercentage = (reductionAmount / contractedDemand) * 100;
        }

        return [
            row.serialNo,
            row.date,
            row.customerName,
            row.billingId,
            row.customerId,
            row.contractedDemand,
            row.address,
            morningLoad,
            eveningLoad,
            reductionAmount.toFixed(2),
            reductionPercentage.toFixed(2),
            ...row.hourlyData
        ];
    });

    const timeHeaders = Array.from({ length: 96 }, (_, i) => {
        const hour = Math.floor(i / 4);
        const minute = (i % 4) * 15;
        return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')} [KW]`;
    });

    const headers = [
        'Serial no.', 'Date', 'Customer name', 'Billing id', 'Customer id', 'Contracted demand', 'Address',
        'Morning Load (KW)', 'Evening Load (KW)', 'Reduction Amount (KW)', 'Reduction Percentage (%)',
        ...timeHeaders
    ];

    const ws = XLSX.utils.aoa_to_sheet([headers, ...exportData]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "CustomerData");
    XLSX.writeFile(wb, "CustomerData.xlsx");
    Swal.fire('موفقیت', 'فایل اکسل با موفقیت ذخیره شد.', 'success');
    log('info', 'Data exported to Excel successfully.');
}

/**
 * تابع خروجی PDF (فقط اطلاعات اصلی مشترک و بارهای محاسبه شده)
 */
async function exportToPdf() {
    if (filteredData.length === 0) {
        Swal.fire('توجه', 'داده‌ای برای خروجی گرفتن وجود ندارد.', 'warning');
        log('warn', 'Attempted to export to PDF with no data.');
        return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({
        orientation: 'landscape', // Landscape for wider tables
        unit: 'pt',
        format: 'a4'
    });

    try {
        // Ensure the path to the font is correct relative to Charts10.html
        const fontBytes = await fetch('fonts/Vazirmatn-Regular.ttf').then(res => res.arrayBuffer());
        doc.addFileToVFS('Vazirmatn-Regular.ttf', btoa(String.fromCharCode(...new Uint8Array(fontBytes))));
        doc.addFont('Vazirmatn-Regular.ttf', 'Vazirmatn', 'normal');
        doc.setFont('Vazirmatn');
        log('info', 'Vazirmatn font loaded successfully for PDF.');
    } catch (error) {
        log('error', 'Failed to load Vazirmatn font for PDF. PDF might not display Persian text correctly.', { errorMessage: error.message });
        Swal.fire('خطا', 'فونت فارسی برای PDF بارگذاری نشد. متن فارسی ممکن است به درستی نمایش داده نشود.', 'warning');
    }

    const tableHeaders = [
        ['ردیف', 'سریال', 'تاریخ', 'نام مشترک', 'شناسه قبض', 'شناسه مشترک', 'دیماند قراردادی (KW)', 'آدرس', 'بار صبح (KW)', 'بار عصر (KW)', 'میزان کاهش (KW)', 'درصد کاهش (%)']
    ];

    const tableBody = filteredData.map((row, index) => {
        const morningPeakStart = document.getElementById('morningPeakStart').value;
        const morningPeakEnd = document.getElementById('morningPeakEnd').value;
        const eveningPeakStart = document.getElementById('eveningPeakStart').value;
        const eveningPeakEnd = document.getElementById('eveningPeakEnd').value;
        const morningCalcType = document.getElementById('morningCalcType').value;
        const eveningCalcType = document.getElementById('eveningCalcType').value;

        const morningLoad = calculatePeakLoad(row.hourlyData, morningPeakStart, morningPeakEnd, morningCalcType).toFixed(2);
        const eveningLoad = calculatePeakLoad(row.hourlyData, eveningPeakStart, eveningPeakEnd, eveningCalcType).toFixed(2);

        const contractedDemand = parseFloat(row.contractedDemand);
        let reductionAmount = 0;
        let reductionPercentage = 0;
        if (contractedDemand > 0) {
            const validHourlyData = row.hourlyData.filter(val => typeof val === 'number' && !isNaN(val));
            const maxLoad = validHourlyData.length > 0 ? Math.max(...validHourlyData) : 0;
            reductionAmount = contractedDemand - maxLoad;
            reductionPercentage = (reductionAmount / contractedDemand) * 100;
        }

        return [
            index + 1, // Row number
            row.serialNo,
            row.date,
            row.customerName,
            row.billingId,
            row.customerId,
            row.contractedDemand,
            row.address,
            morningLoad,
            eveningLoad,
            reductionAmount.toFixed(2),
            reductionPercentage.toFixed(2) + '%'
        ];
    });

    doc.autoTable({
        head: tableHeaders,
        body: tableBody,
        startY: 60, // Start table below header
        theme: 'striped',
        styles: {
            font: 'Vazirmatn',
            fontStyle: 'normal',
            fontSize: 6, // Reduced font size for more compactness
            cellPadding: 0.5, // Reduced padding
            halign: 'center', // Center align text
            overflow: 'linebreak', // Allow text to wrap
            cellWidth: 'auto' // Let autoTable decide optimal width
        },
        headStyles: {
            fillColor: [46, 204, 113], // Green header
            textColor: [255, 255, 255],
            fontSize: 7, // Slightly larger header font
            cellPadding: 1, // Slightly larger header padding
            font: 'Vazirmatn',
        },
        bodyStyles: {
            textColor: [51, 51, 51],
        },
        alternateRowStyles: {
            fillColor: [240, 240, 240]
        },
        columnStyles: { // Specific column styles if needed, adjust as per your data length
            // 0: {cellWidth: 20}, // Row #
            // 1: {cellWidth: 40}, // Serial no.
            // 2: {cellWidth: 50}, // Date
            // 3: {cellWidth: 80}, // Customer name
            // 4: {cellWidth: 50}, // Billing id
            // 5: {cellWidth: 50}, // Customer id
            // 6: {cellWidth: 60}, // Contracted demand
            // 7: {cellWidth: 100},// Address
            // 8: {cellWidth: 50}, // Morning Load
            // 9: {cellWidth: 50}, // Evening Load
            // 10: {cellWidth: 50},// Reduction Amount
            // 11: {cellWidth: 50} // Reduction Percentage
        }
    });

    doc.save("CustomerData_Summary.pdf");
    Swal.fire('موفقیت', 'فایل PDF (جداول) با موفقیت ذخیره شد.', 'success');
    log('info', 'Data exported to PDF (tables) successfully.');
}

/**
 * تابع خروجی گرفتن از نمودارها به عنوان PDF (تصویر)
 */
async function exportChartsAsImages() {
    if (currentCharts.length === 0) {
        Swal.fire('توجه', 'هیچ نموداری برای خروجی گرفتن وجود ندارد. ابتدا نمودارها را نمایش دهید (با کلیک بر روی دکمه نمودار کنار هر ردیف یا دکمه "رسم همه نمودارها").', 'warning');
        log('warn', 'Attempted to export charts as images with no charts available.');
        return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({
        orientation: 'portrait', // Portrait is often better for single charts
        unit: 'pt',
        format: 'a4'
    });

    let yOffset = 20;
    const margin = 10;
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();

    showProgress(0, 'در حال تولید PDF نمودارها...');
    for (let i = 0; i < currentCharts.length; i++) {
        const chart = currentCharts[i];
        const canvas = chart.canvas;

        try {
            const imgData = await html2canvas(canvas, { scale: 2, backgroundColor: '#ffffff' }).then(canvas => canvas.toDataURL('image/png'));
            const imgProps = doc.getImageProperties(imgData);
            const imgWidth = pageWidth - 2 * margin;
            const imgHeight = (imgProps.height * imgWidth) / imgProps.width;

            if (yOffset + imgHeight + margin > pageHeight && i > 0) { // Add page if current chart doesn't fit, but not for the very first chart
                doc.addPage();
                yOffset = 20;
                log('info', `Added new page for charts during PDF export.`);
            }

            doc.addImage(imgData, 'PNG', margin, yOffset, imgWidth, imgHeight);
            yOffset += imgHeight + margin;
            log('info', `Chart ${i + 1} added to PDF.`);
            showProgress(Math.round(((i + 1) / currentCharts.length) * 100), `در حال اضافه کردن نمودار ${i + 1} از ${currentCharts.length} به PDF`);

        } catch (error) {
            log('error', `Failed to convert chart ${i + 1} to image for PDF.`, { errorMessage: error.message, chartIndex: i });
            Swal.fire('خطا', `خطا در تولید تصویر برای نمودار ${i + 1}: ${error.message}`, 'error');
            showProgress(0, 'خطا در خروجی PDF نمودارها');
            return;
        }
    }

    doc.save("CustomerCharts.pdf");
    Swal.fire('موفقیت', 'نمودارها با موفقیت به صورت PDF ذخیره شدند.', 'success');
    showProgress(100, 'خروجی PDF نمودارها کامل شد.');
    log('info', 'Charts exported as PDF successfully.');
}

/**
 * تابع دانلود فایل لاگ برنامه
 */
function downloadLogFile() {
    if (appLogs.length === 0) {
        Swal.fire('توجه', 'هیچ لاگی برای دانلود وجود ندارد.', 'info');
        return;
    }

    const logContent = appLogs.map(entry => {
        let logString = `[${entry.timestamp}] [${entry.level}]: ${entry.message}`;
        if (Object.keys(entry.context).length > 0) {
            logString += ` - Context: ${JSON.stringify(entry.context)}`;
        }
        return logString;
    }).join('\n');

    const blob = new Blob([logContent], { type: "text/plain;charset=utf-8" });
    saveAs(blob, `application_log_${new Date().toISOString().split('T')[0]}.txt`);
    log('info', 'Application logs downloaded.');
}


// ====================================================================================================
// مدیریت رویدادها (Event Listeners)
// ====================================================================================================

document.addEventListener('DOMContentLoaded', () => {
    log('info', 'DOM fully loaded and parsed.');

    // Get DOM elements
    const fileInput = document.getElementById('excelFile');
    const fileNameDisplay = document.getElementById('fileNameDisplay');
    const sheetSelect = document.getElementById('sheetSelect');
    const processDataBtn = document.getElementById('processDataBtn');
    const resetAppBtn = document.getElementById('resetAppBtn');
    const resultsTableBody = document.querySelector('#resultsTable tbody');
    const exportExcelBtn = document.getElementById('exportExcelBtn');
    const exportPdfBtn = document.getElementById('exportPdfBtn');
    const exportChartsAsImagesBtn = document.getElementById('exportChartsAsImagesBtn');
    const exportLogFileBtn = document.getElementById('exportLogFile');
    const applyFilterBtn = document.getElementById('applyFilterBtn');
    const clearFilterBtn = document.getElementById('clearFilterBtn');
    // Removed specific checkboxes for evening/reduction, as per understanding of your preferred simplified UI
    // const chkEvening = document.getElementById('chkEvening');
    // const txtEvening = document.getElementById('txtEvening');
    // const chkReduction = document.getElementById('chkReduction');
    // const txtReduction = document.getElementById('txtReduction');
    const processStatusMessage = document.getElementById('processStatusMessage');
    const renderAllChartsBtn = document.getElementById('renderAllChartsBtn'); // New button

    // Initial state setup
    showProgress(0, 'منتظر انتخاب فایل...');
    if (fileNameDisplay) fileNameDisplay.textContent = 'فایل انتخاب نشده است.';
    if (processDataBtn) processDataBtn.disabled = true;
    if (sheetSelect) sheetSelect.disabled = true;
    if (exportExcelBtn) exportExcelBtn.disabled = true;
    if (exportPdfBtn) exportPdfBtn.disabled = true;
    if (exportChartsAsImagesBtn) exportChartsAsImagesBtn.disabled = true;
    if (renderAllChartsBtn) renderAllChartsBtn.disabled = true; // Disable new button initially
    // if (txtEvening) txtEvening.disabled = true; // No longer needed
    // if (txtReduction) txtReduction.disabled = true; // No longer needed
    if (processStatusMessage) processStatusMessage.textContent = ''; // Clear initial status message

    // Event listeners

    // File input change
    if (fileInput) {
        fileInput.addEventListener('change', () => {
            if (fileInput.files && fileInput.files[0]) {
                if (fileNameDisplay) fileNameDisplay.textContent = fileInput.files[0].name;
                if (processDataBtn) processDataBtn.disabled = false;
                if (processStatusMessage) processStatusMessage.textContent = 'فایل آماده پردازش است.';
                log('info', `File selected: ${fileInput.files[0].name}`);
            } else {
                if (fileNameDisplay) fileNameDisplay.textContent = 'فایل انتخاب نشده است.';
                if (processDataBtn) processDataBtn.disabled = true;
                if (processStatusMessage) processStatusMessage.textContent = '';
                log('info', 'No file selected.');
            }
        });
    }

    // Sheet select change (re-process data for the selected sheet)
    if (sheetSelect) {
        sheetSelect.addEventListener('change', () => {
            if (workbook && sheetSelect.value) {
                processCurrentSheet(sheetSelect.value);
                log('info', `Sheet changed to: ${sheetSelect.value}. Re-processing data.`);
            } else {
                log('warn', 'Workbook not loaded or sheet not selected when attempting to change sheet.');
            }
        });
    }

    // Process data button click
    if (processDataBtn) {
        processDataBtn.addEventListener('click', processExcelFile);
    }

    // Reset app button click
    if (resetAppBtn) {
        resetAppBtn.addEventListener('click', () => {
            workbook = null;
            parsedData = [];
            filteredData = [];
            currentCharts = [];
            appLogs = []; // Clear logs on reset

            if (fileInput) {
                fileInput.value = '';
                if (fileNameDisplay) fileNameDisplay.textContent = 'فایل انتخاب نشده است.';
            }
            if (sheetSelect) {
                sheetSelect.innerHTML = '<option value="">- Sheet1 -</option>';
                sheetSelect.disabled = true;
            }
            if (processDataBtn) processDataBtn.disabled = true;
            if (exportExcelBtn) exportExcelBtn.disabled = true;
            if (exportPdfBtn) exportPdfBtn.disabled = true;
            if (exportChartsAsImagesBtn) exportChartsAsImagesBtn.disabled = true;
            if (renderAllChartsBtn) renderAllChartsBtn.disabled = true; // Disable new button on reset
            if (resultsTableBody) resultsTableBody.innerHTML = '';
            destroyCharts();
            if (document.getElementById('noChartsMessage')) document.getElementById('noChartsMessage').style.display = 'block';

            // if (chkEvening) chkEvening.checked = false; // No longer needed
            // if (txtEvening) txtEvening.disabled = true; // No longer needed
            // if (chkReduction) chkReduction.checked = false; // No longer needed
            // if (txtReduction) txtReduction.disabled = true; // No longer needed
            if (document.getElementById('filterValue')) document.getElementById('filterValue').value = '';
            if (document.getElementById('filterColumn')) document.getElementById('filterColumn').value = 'all';

            showProgress(0, 'منتظر انتخاب فایل...');
            if (processStatusMessage) processStatusMessage.textContent = '';
            Swal.fire('با موفقیت', 'برنامه به حالت اولیه بازنشانی شد.', 'success');
            log('info', 'Application reset successfully.');
        });
    }

    // Checkbox listeners for evening/reduction settings - Removed as per current UI
    // if (chkEvening) {
    //     chkEvening.addEventListener('change', () => {
    //         if (txtEvening) txtEvening.disabled = !chkEvening.checked;
    //     });
    // }
    // if (chkReduction) {
    //     chkReduction.addEventListener('change', () => {
    //         if (txtReduction) txtReduction.disabled = !chkReduction.checked;
    //     });
    // }

    // Filter buttons
    if (applyFilterBtn) {
        applyFilterBtn.addEventListener('click', applyFilter);
    }
    if (clearFilterBtn) {
        clearFilterBtn.addEventListener('click', clearFilter);
    }

    // Export buttons
    if (exportExcelBtn) {processExcel
        exportExcelBtn.addEventListener('click', exportToExcel);
    }
    if (exportPdfBtn) {
        exportPdfBtn.addEventListener('click', exportToPdf);
    }
    if (exportChartsAsImagesBtn) {
        exportChartsAsImagesBtn.addEventListener('click', exportChartsAsImages);
    }
    if (exportLogFileBtn) {
        exportLogFileBtn.addEventListener('click', downloadLogFile);
    }

    // New "Render All Charts" button
    if (renderAllChartsBtn) {
        renderAllChartsBtn.addEventListener('click', renderAllCharts);
    }
});