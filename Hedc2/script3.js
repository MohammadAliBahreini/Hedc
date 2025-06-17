// script.js

// Global variables to store parsed data and chart instances
// Declaring them once at the top level to avoid "already declared" errors.
let workbook = null; // Stores the Excel workbook data
let parsedData = []; // Stores the processed data for the table
let charts = []; // Stores Chart.js instances

// Register Chart.js plugins here, after global variables but before DOM access.
// Ensure Chart and ChartjsPluginAnnotation are loaded via <script> tags in HTML first.
try {
    if (typeof Chart !== 'undefined' && typeof ChartjsPluginAnnotation !== 'undefined') {
        Chart.register(ChartjsPluginAnnotation);
        console.log("پلاگین ChartjsPluginAnnotation با موفقیت ثبت شد.");
    } else {
        console.error("خطای حیاتی: Chart یا ChartjsPluginAnnotation تعریف نشده‌اند. نمودارها نمایش داده نمی‌شوند.");
        Swal.fire({
            icon: 'error',
            title: 'خطای بارگذاری',
            text: 'برخی از قابلیت‌های برنامه به دلیل مشکل در بارگذاری کتابخانه‌ها ممکن است کار نکنند. لطفاً با پشتیبانی تماس بگیرید.'
        });
    }
} catch (e) {
    console.error("خطا هنگام Chart.register:", e);
    Swal.fire({
        icon: 'error',
        title: 'خطای ثبت پلاگین',
        text: 'مشکلی در ثبت پلاگین‌های نمودار وجود دارد. لطفاً برنامه را مجدداً راه‌اندازی کنید.'
    });
}


// --- DOM Element References ---
const excelFile = document.getElementById('excelFile');
const fileNameDisplay = document.getElementById('fileNameDisplay');
const sheetSelect = document.getElementById('sheetSelect');
const processDataBtn = document.getElementById('processDataBtn');
const resetAppBtn = document.getElementById('resetAppBtn');
const resultsTableBody = document.querySelector('#resultsTable tbody');
const chartsContainer = document.getElementById('chartsContainer');
const noChartsMessage = document.getElementById('noChartsMessage');
const exportExcelBtn = document.getElementById('exportExcelBtn');
const exportPdfBtn = document.getElementById('exportPdfBtn');
const progressContainer = document.getElementById('progress-container');
const progressBar = document.getElementById('progress-bar');
const progressLabel = document.getElementById('progress-label');

// Filter and Calculation Inputs
const chkEvening = document.getElementById('chkEvening');
const txtEvening = document.getElementById('txtEvening');
const chkReduction = document.getElementById('chkReduction');
const txtReduction = document.getElementById('txtReduction');

const morningCalcType = document.getElementById('morningCalcType');
const morningStartHour = document.getElementById('morningStartHour');
const morningStartMinute = document.getElementById('morningStartMinute');
const morningEndHour = document.getElementById('morningEndHour');
const morningEndMinute = document.getElementById('morningEndMinute');

const eveningCalcType = document.getElementById('eveningCalcType');
const eveningStartHour = document.getElementById('eveningStartHour');
const eveningStartMinute = document.getElementById('eveningStartMinute');
const eveningEndHour = document.getElementById('eveningEndHour');
const eveningEndMinute = document.getElementById('eveningEndMinute');

// --- Event Listeners ---

// Handle file selection
excelFile.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (file) {
        fileNameDisplay.textContent = file.name;
        progressLabel.textContent = 'فایل در حال بارگذاری...';
        showProgress(10); // Initial progress

        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                workbook = XLSX.read(data, { type: 'array' });

                // Populate sheet selection
                sheetSelect.innerHTML = '';
                workbook.SheetNames.forEach(sheetName => {
                    const option = document.createElement('option');
                    option.value = sheetName;
                    option.textContent = sheetName;
                    sheetSelect.appendChild(option);
                });
                sheetSelect.disabled = false;
                processDataBtn.disabled = false;
                progressLabel.textContent = 'فایل بارگذاری شد. شیت را انتخاب کنید.';
                showProgress(100);
            } catch (error) {
                console.error("خطا در خواندن فایل اکسل:", error);
                Swal.fire({
                    icon: 'error',
                    title: 'خطا در بارگذاری فایل',
                    text: 'فایل اکسل نامعتبر است یا در خواندن آن مشکلی پیش آمد.'
                });
                resetApplication();
            }
        };
        reader.onerror = (e) => {
            console.error("خطا در FileReader:", e);
            Swal.fire({
                icon: 'error',
                title: 'خطا در خواندن فایل',
                text: 'مشکلی در خواندن فایل پیش آمد. لطفاً دوباره امتحان کنید.'
            });
            resetApplication();
        };
        reader.readAsArrayBuffer(file);
    } else {
        fileNameDisplay.textContent = 'فایل انتخاب نشده...';
        sheetSelect.innerHTML = '<option value="">- Sheet1 -</option>';
        sheetSelect.disabled = true;
        processDataBtn.disabled = true;
        progressLabel.textContent = 'منتظر انتخاب فایل...';
        showProgress(0);
        resetApplication();
    }
});

// Handle sheet selection (optional, can trigger re-processing if desired)
sheetSelect.addEventListener('change', () => {
    // If you want to automatically re-process on sheet change, uncomment:
    // if (workbook) {
    //     processData();
    // }
});

// Process data button click
processDataBtn.addEventListener('click', processData);

// Reset application button
resetAppBtn.addEventListener('click', () => {
    Swal.fire({
        title: 'اطمینان دارید؟',
        text: "تمام داده‌ها و نمودارها پاک خواهند شد!",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#3085d6',
        cancelButtonColor: '#d33',
        confirmButtonText: 'بله، شروع مجدد',
        cancelButtonText: 'خیر'
    }).then((result) => {
        if (result.isConfirmed) {
            resetApplication();
            Swal.fire(
                'شروع مجدد!',
                'برنامه بازنشانی شد.',
                'success'
            );
        }
    });
});

// Filter checkbox event listeners
chkEvening.addEventListener('change', () => {
    txtEvening.disabled = !chkEvening.checked;
    if (parsedData.length > 0) processData(); // Re-process if data exists
});
chkReduction.addEventListener('change', () => {
    txtReduction.disabled = !chkReduction.checked;
    if (parsedData.length > 0) processData(); // Re-process if data exists
});

// Time input change listeners (debounce or processData on change)
[
    morningCalcType, morningStartHour, morningStartMinute,
    morningEndHour, morningEndMinute,
    eveningCalcType, eveningStartHour, eveningStartMinute,
    eveningEndHour, eveningEndMinute
].forEach(input => {
    input.addEventListener('change', () => {
        if (parsedData.length > 0) {
            processData(); // Re-process if data exists
        }
    });
});

// Export buttons
exportExcelBtn.addEventListener('click', exportToExcel);
exportPdfBtn.addEventListener('click', exportToPdf);

// --- Core Functions ---

function showProgress(percentage, label = null) {
    progressBar.style.width = percentage + '%';
    progressBar.setAttribute('aria-valuenow', percentage);
    if (label) {
        progressLabel.textContent = label;
    }
    progressContainer.style.display = percentage > 0 ? 'block' : 'none';
}

function resetApplication() {
    excelFile.value = '';
    fileNameDisplay.textContent = 'فایل انتخاب نشده...';
    sheetSelect.innerHTML = '<option value="">- Sheet1 -</option>';
    sheetSelect.disabled = true;
    processDataBtn.disabled = true;
    exportExcelBtn.disabled = true;
    exportPdfBtn.disabled = true;
    resultsTableBody.innerHTML = '';
    chartsContainer.innerHTML = '<p id="noChartsMessage" style="text-align: center; color: #777; margin-top: 20px;">برای نمایش نمودارها، لطفاً ابتدا داده‌ها را پردازش کنید.</p>';
    noChartsMessage.style.display = 'block'; // Ensure message is visible

    workbook = null;
    parsedData = [];
    destroyCharts(); // Clear existing Chart.js instances

    progressLabel.textContent = 'منتظر انتخاب فایل...';
    showProgress(0);

    // Reset filter inputs to default enabled/disabled state based on checkboxes
    txtEvening.disabled = !chkEvening.checked;
    txtReduction.disabled = !chkReduction.checked;
}

function destroyCharts() {
    charts.forEach(chart => {
        if (chart) {
            chart.destroy();
        }
    });
    charts = []; // Clear the array
}

async function processData() {
    if (!workbook) {
        Swal.fire('خطا', 'لطفاً ابتدا یک فایل اکسل انتخاب کنید.', 'error');
        return;
    }

    const selectedSheetName = sheetSelect.value || workbook.SheetNames[0];
    if (!selectedSheetName) {
        Swal.fire('خطا', 'شیتی برای پردازش یافت نشد.', 'error');
        return;
    }

    progressLabel.textContent = 'در حال پردازش داده‌ها...';
    showProgress(25);

    // Clear previous results and charts
    resultsTableBody.innerHTML = '';
    destroyCharts(); // Destroy existing chart instances
    chartsContainer.innerHTML = ''; // Clear chart container
    noChartsMessage.style.display = 'none'; // Hide the no charts message

    // Read data from the selected sheet
    const worksheet = workbook.Sheets[selectedSheetName];
    // Convert sheet to JSON, skipping header (header:1) to get array of arrays
    // Use raw: true to get raw values, not formatted ones
    const jsonSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });

    // Assuming the first row is headers and data starts from the second row
    const headers = jsonSheet[0];
    // console.log("Headers from Excel:", headers);
    const dataRows = jsonSheet.slice(1);
    console.log(parsedData)
    console.log("Data Rows:", dataRows);
    parsedData = []; // Clear previous data
    // Removed timeColumnIndex as it was unused and might cause confusion.
    // Time parsing will now happen from headers themselves.

    // Parse time boundaries once
    const getMinutes = (h, m) => parseInt(h) * 60 + parseInt(m);
    const morningStart = getMinutes(morningStartHour.value, morningStartMinute.value);
    const morningEnd = getMinutes(morningEndHour.value, morningEndMinute.value);
    const eveningStart = getMinutes(eveningStartHour.value, eveningStartMinute.value);
    const eveningEnd = getMinutes(eveningEndHour.value, eveningEndMinute.value);

    // Identify relevant columns by header name
    const bodyNumberColIndex = headers.indexOf('Serial no.');
    const customerNameColIndex = headers.indexOf('Customer name');
    const billIdColIndex = headers.indexOf('Billing id');
    const addressColIndex = headers.indexOf('Address');
    const subscriptionNumberColIndex = headers.indexOf('Customer id');
    const demandColumnIndex = headers.indexOf('Contracted demand');

    // Check if essential columns exist
    if (bodyNumberColIndex === -1 || customerNameColIndex === -1 || billIdColIndex === -1 ||
        addressColIndex === -1 || subscriptionNumberColIndex === -1 || demandColumnIndex === -1) {
        Swal.fire({
            icon: 'error',
            title: 'خطای ساختار فایل',
            html: 'یکی از ستون‌های ضروری (مانند "شماره بدنه", "نام مشترک", "شناسه قبض", "آدرس مشترک", "شماره اشتراک", "دیماند قراردادی (KW)") در فایل اکسل یافت نشد. <br> لطفاً مطمئن شوید که سربرگ‌ها صحیح هستند.'
        });
        showProgress(0);
        return;
    }


    // Filter values
    const minEveningLoad = chkEvening.checked ? parseFloat(txtEvening.value) : -Infinity;
    const minReductionPercent = chkReduction.checked ? parseFloat(txtReduction.value) : -Infinity;

    progressLabel.textContent = 'در حال تحلیل بارهای مشترکین...';
    showProgress(50);

    for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
        // Skip empty rows or rows that don't look like data rows
        if (!row || row.length === 0 || !row[bodyNumberColIndex]) continue;

        const customerInfo = {
            rowNum: i + 2, // Original row number in Excel (assuming header is row 1)
            bodyNumber: row[bodyNumberColIndex],
            customerName: row[customerNameColIndex],
            billId: row[billIdColIndex],
            address: row[addressColIndex],
            subscriptionNumber: row[subscriptionNumberColIndex],
            contractDemand: parseFloat(row[demandColumnIndex]) || 0 // Default to 0 if invalid
        };

        const loadProfile = []; // To store { timeInMinutes, loadValue } for each row
        const timeLabels = []; // To store time labels for chart

        // Loop through all columns starting from the first potential load data column.
        // A more robust way might be to identify load columns by a pattern (e.g., "00:00")
        // For now, assuming load data starts after the fixed info columns.
        // If your first 6 columns are info, load data starts from index 6.
        const firstLoadColumnIndex = Math.max(
            bodyNumberColIndex, customerNameColIndex, billIdColIndex,
            addressColIndex, subscriptionNumberColIndex, demandColumnIndex
        ) + 1; // Start reading from the column right after the last info column

        for (let j = firstLoadColumnIndex; j < headers.length; j++) {
            const header = headers[j];
            // Check if the header looks like a time (e.g., "00:00", "00:15")
            if (typeof header === 'string' && header.match(/^\d{2}:\d{2}$/)) {
                const [h, m] = header.split(':').map(Number);
                const timeInMinutes = h * 60 + m;
                const loadValue = parseFloat(row[j]);
                if (!isNaN(loadValue)) {
                    loadProfile.push({ timeInMinutes, load: loadValue });
                    timeLabels.push(header); // Use original string for label
                }
            }
        }

        if (loadProfile.length === 0) {
            console.log("Skipping row due to empty load profile:", customerInfo.bodyNumber);
            continue; // Skip if no load data found for this customer
        }

        // Calculate morning load
        const morningLoads = loadProfile.filter(item =>
            item.timeInMinutes >= morningStart && item.timeInMinutes <= morningEnd
        ).map(item => item.load);

        let morningLoad = 0;
        if (morningLoads.length > 0) {
            if (morningCalcType.value === 'avg') morningLoad = morningLoads.reduce((a, b) => a + b, 0) / morningLoads.length;
            else if (morningCalcType.value === 'max') morningLoad = Math.max(...morningLoads);
            else if (morningCalcType.value === 'min') morningLoad = Math.min(...morningLoads);
        }

        // Calculate evening load
        const eveningLoads = loadProfile.filter(item =>
            item.timeInMinutes >= eveningStart && item.timeInMinutes <= eveningEnd
        ).map(item => item.load);

        let eveningLoad = 0;
        if (eveningLoads.length > 0) {
            if (eveningCalcType.value === 'avg') eveningLoad = eveningLoads.reduce((a, b) => a + b, 0) / eveningLoads.length;
            else if (eveningCalcType.value === 'max') eveningLoad = Math.max(...eveningLoads);
            else if (eveningCalcType.value === 'min') eveningLoad = Math.min(...eveningLoads);
        }

        const reductionKW = morningLoad - eveningLoad;
        const reductionPercent = (morningLoad > 0) ? (reductionKW / morningLoad) * 100 : 0;

        const customerResult = {
            ...customerInfo,
            morningLoad: morningLoad.toFixed(2),
            eveningLoad: eveningLoad.toFixed(2),
            reductionKW: reductionKW.toFixed(2),
            reductionPercent: reductionPercent.toFixed(2),
            loadProfileData: loadProfile.map(item => item.load), // Only load values for chart data
            timeLabels: timeLabels // Time labels for chart X-axis
        };

        // Apply filters
        const passesEveningFilter = !chkEvening.checked || (parseFloat(customerResult.eveningLoad) >= minEveningLoad);
        const passesReductionFilter = !chkReduction.checked || (parseFloat(customerResult.reductionPercent) >= minReductionPercent);

        if (passesEveningFilter && passesReductionFilter) {
            parsedData.push(customerResult);
            console.log("Customer added:", customerResult);
        }
    }

    showProgress(75, 'در حال نمایش نتایج...');
    displayResults();
    drawCharts();
    showProgress(100, 'پردازش کامل شد.');

    exportExcelBtn.disabled = false;
    exportPdfBtn.disabled = false;
}


function displayResults() {
    resultsTableBody.innerHTML = ''; // Clear existing rows
    if (parsedData.length === 0) {
        const noDataRow = resultsTableBody.insertRow();
        const cell = noDataRow.insertCell();
        cell.colSpan = 11; // Span across all columns
        cell.textContent = 'هیچ داده‌ای مطابق با فیلترها یافت نشد.';
        cell.style.textAlign = 'center';
        exportExcelBtn.disabled = true;
        exportPdfBtn.disabled = true;
        chartsContainer.innerHTML = '<p id="noChartsMessage" style="text-align: center; color: #777; margin-top: 20px;">برای نمایش نمودارها، لطفاً ابتدا داده‌ها را پردازش کنید.</p>';
        noChartsMessage.style.display = 'block'; // Ensure message is visible
        return;
    }

    parsedData.forEach((data, index) => {
        const row = resultsTableBody.insertRow();
        row.insertCell().textContent = index + 1; // Rownum in the results table
        row.insertCell().textContent = data.bodyNumber;
        row.insertCell().textContent = data.customerName;
        row.insertCell().textContent = data.billId;
        row.insertCell().textContent = data.address;
        row.insertCell().textContent = data.subscriptionNumber;
        row.insertCell().textContent = data.contractDemand;
        row.insertCell().textContent = data.morningLoad;
        row.insertCell().textContent = data.eveningLoad;
        row.insertCell().textContent = data.reductionKW;
        row.insertCell().textContent = data.reductionPercent;
    });
}

function drawCharts() {
    destroyCharts(); // Clear any previous chart instances
    chartsContainer.innerHTML = ''; // Clear container
    if (parsedData.length === 0) {
        noChartsMessage.style.display = 'block'; // Show the no charts message
        return;
    }
    noChartsMessage.style.display = 'none'; // Hide the no charts message if data exists

    parsedData.forEach((customer, index) => {
        const chartWrapper = document.createElement('div');
        chartWrapper.className = 'chart-wrapper';
        chartWrapper.id = `chart-wrapper-${index}`; // Unique ID for each chart wrapper

        const canvas = document.createElement('canvas');
        canvas.id = `chart-${index}`; // Unique ID for each canvas
        chartWrapper.appendChild(canvas);

        const downloadBtn = document.createElement('button');
        downloadBtn.className = 'chart-download-btn';
        downloadBtn.textContent = 'دانلود نمودار (PNG)';
        downloadBtn.addEventListener('click', () => downloadChartAsPNG(canvas.id, customer.customerName, customer.bodyNumber));
        chartWrapper.appendChild(downloadBtn);

        chartsContainer.appendChild(chartWrapper);

        const ctx = canvas.getContext('2d');

        const chartData = {
            labels: customer.timeLabels,
            datasets: [{
                label: `مصرف برق (KW)`,
                data: customer.loadProfileData,
                borderColor: 'rgb(75, 192, 192)',
                backgroundColor: 'rgba(75, 192, 192, 0.2)',
                borderWidth: 2,
                pointRadius: 0,
                fill: false
            }]
        };

        // Prepare annotations
        // Ensure xMin/xMax are valid indices within timeLabels
        const morningStartIndex = customer.timeLabels.indexOf(`${String(morningStartHour.value).padStart(2, '0')}:${String(morningStartMinute.value).padStart(2, '0')}`);
        const morningEndIndex = customer.timeLabels.indexOf(`${String(morningEndHour.value).padStart(2, '0')}:${String(morningEndMinute.value).padStart(2, '0')}`);
        const eveningStartIndex = customer.timeLabels.indexOf(`${String(eveningStartHour.value).padStart(2, '0')}:${String(eveningStartMinute.value).padStart(2, '0')}`);
        const eveningEndIndex = customer.timeLabels.indexOf(`${String(eveningEndHour.value).padStart(2, '0')}:${String(eveningEndMinute.value).padStart(2, '0')}`);

        const annotations = {};

        if (morningStartIndex !== -1 && morningEndIndex !== -1) {
            annotations.morningPeak = {
                type: 'box',
                xMin: morningStartIndex,
                xMax: morningEndIndex,
                backgroundColor: 'rgba(255, 99, 132, 0.2)',
                borderColor: 'rgb(255, 99, 132)',
                borderWidth: 1,
                label: {
                    content: 'ساعات اوج بار صبح',
                    enabled: true,
                    position: 'top',
                    backgroundColor: 'rgba(255, 99, 132, 0.7)',
                    color: '#fff',
                    font: { family: 'Vazirmatn', size: 10 }
                }
            };
        }

        if (eveningStartIndex !== -1 && eveningEndIndex !== -1) {
            annotations.eveningPeak = {
                type: 'box',
                xMin: eveningStartIndex,
                xMax: eveningEndIndex,
                backgroundColor: 'rgba(54, 162, 235, 0.2)',
                borderColor: 'rgb(54, 162, 235)',
                borderWidth: 1,
                label: {
                    content: 'ساعات اوج بار عصر',
                    enabled: true,
                    position: 'top',
                    backgroundColor: 'rgba(54, 162, 235, 0.7)',
                    color: '#fff',
                    font: { family: 'Vazirmatn', size: 10 }
                }
            };
        }

        if (customer.contractDemand > 0) {
            annotations.contractDemandLine = {
                type: 'line',
                yMin: customer.contractDemand,
                yMax: customer.contractDemand,
                borderColor: 'rgb(255, 159, 64)',
                borderWidth: 2,
                borderDash: [6, 6],
                label: {
                    content: `دیماند قراردادی: ${customer.contractDemand} KW`,
                    enabled: true,
                    position: 'end',
                    backgroundColor: 'rgba(255, 159, 64, 0.7)',
                    color: '#fff',
                    font: { family: 'Vazirmatn', size: 10 }
                }
            };
        }

        const newChart = new Chart(ctx, {
            type: 'line',
            data: chartData,
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    tooltip: {
                        mode: 'index',
                        intersect: false,
                        rtl: true,
                        callbacks: {
                            title: function(context) {
                                return `زمان: ${context[0].label}`;
                            },
                            label: function(context) {
                                let label = context.dataset.label || '';
                                if (label) {
                                    label += ': ';
                                }
                                if (context.parsed.y !== null) {
                                    label += new Intl.NumberFormat('fa-IR', { maximumFractionDigits: 2 }).format(context.parsed.y) + ' KW';
                                }
                                return label;
                            }
                        },
                        titleFont: { family: 'Vazirmatn' },
                        bodyFont: { family: 'Vazirmatn' }
                    },
                    legend: {
                        display: true,
                        position: 'top',
                        rtl: true,
                        labels: {
                            font: {
                                family: 'Vazirmatn',
                                size: 12
                            }
                        }
                    },
                    title: {
                        display: true,
                        text: `پروفیل بار مشترک: ${customer.customerName} (${customer.bodyNumber}) - دیماند قراردادی: ${customer.contractDemand} KW`,
                        font: {
                            size: 14,
                            family: 'Vazirmatn',
                            weight: 'bold'
                        },
                        rtl: true
                    },
                    annotation: {
                        annotations: annotations
                    }
                },
                scales: {
                    x: {
                        title: {
                            display: true,
                            text: 'زمان (ساعت:دقیقه)',
                            font: {
                                family: 'Vazirmatn',
                                size: 12
                            }
                        },
                        ticks: {
                            autoSkip: true,
                            maxTicksLimit: 20,
                            font: {
                                family: 'Vazirmatn'
                            }
                        },
                        rtl: true
                    },
                    y: {
                        title: {
                            display: true,
                            text: 'مصرف برق (KW)',
                            font: {
                                family: 'Vazirmatn',
                                size: 12
                            }
                        },
                        beginAtZero: true,
                        ticks: {
                            callback: function(value) {
                                return new Intl.NumberFormat('fa-IR', { maximumFractionDigits: 0 }).format(value) + ' KW';
                            },
                            font: {
                                family: 'Vazirmatn'
                            }
                        },
                        rtl: true
                    }
                }
            }
        });
        charts.push(newChart); // Store the chart instance
    });
}

function downloadChartAsPNG(canvasId, customerName, bodyNumber) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) {
        Swal.fire('خطا', 'نمودار پیدا نشد.', 'error');
        return;
    }
    const link = document.createElement('a');
    link.download = `پروفیل_بار_${customerName}_${bodyNumber}.png`;
    link.href = canvas.toDataURL('image/png');
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}

async function exportToExcel() {
    if (parsedData.length === 0) {
        Swal.fire('هیچ داده‌ای برای خروجی اکسل وجود ندارد.', '', 'info');
        return;
    }

    progressLabel.textContent = 'در حال ساخت فایل اکسل...';
    showProgress(10);

    const dataForExcel = parsedData.map(item => ({
        'ردیف': item.rowNum, // Original Excel row number
        'شماره بدنه': item.bodyNumber,
        'نام مشترک': item.customerName,
        'شناسه قبض': item.billId,
        'آدرس مشترک': item.address,
        'شماره اشتراک': item.subscriptionNumber,
        'دیماند قراردادی (KW)': item.contractDemand,
        'بار صبح (KW)': item.morningLoad,
        'بار عصر (KW)': item.eveningLoad,
        'میزان کاهش (KW)': item.reductionKW,
        'درصد کاهش (%)': item.reductionPercent
    }));

    const ws = XLSX.utils.json_to_sheet(dataForExcel);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'نتایج تحلیل مصرف');

    progressLabel.textContent = 'در حال دانلود فایل اکسل...';
    showProgress(70);

    XLSX.writeFile(wb, 'نتایج_تحلیل_مصرف_برق.xlsx');

    showProgress(100, 'فایل اکسل آماده شد.');
    Swal.fire('موفقیت', 'فایل اکسل با موفقیت دانلود شد.', 'success');
    showProgress(0); // Hide progress bar after completion
}

async function exportToPdf() {
    if (parsedData.length === 0) {
        Swal.fire('هیچ نموداری برای خروجی PDF وجود ندارد.', '', 'info');
        return;
    }

    Swal.fire({
        title: 'ساخت PDF',
        text: 'این فرآیند ممکن است کمی طول بکشد، لطفاً منتظر بمانید...',
        allowOutsideClick: false,
        didOpen: () => {
            Swal.showLoading();
        }
    });
    progressLabel.textContent = 'در حال ساخت PDF پروفیل‌ها...';
    showProgress(10);

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'mm', 'a4');

    let yOffset = 10;
    const margin = 10;
    const pageHeight = doc.internal.pageSize.height;
    const pageWidth = doc.internal.pageSize.width;

    // A standard font that supports some Unicode (like Helvetica) might not render Persian correctly.
    // For proper Persian rendering, you need to embed a font like Vazirmatn.
    // This requires converting the font to a base64 string and using doc.addFont.
    // Example:
    // doc.addFont('Vazirmatn-Regular.ttf', 'Vazirmatn', 'normal');
    // doc.setFont('Vazirmatn');
    // For now, using default font with RTL support in text rendering.

    // Calculate height for each chart in PDF (approximate, adjust as needed)
    const chartRenderHeight = 90; // Height of each chart image in mm in PDF
    const chartRenderWidth = pageWidth - (2 * margin); // Max width to fit page with margins

    for (let i = 0; i < charts.length; i++) {
        const chartWrapper = document.getElementById(`chart-wrapper-${i}`); // Get the wrapper
        if (!chartWrapper) continue;

        const customerName = parsedData[i].customerName;
        const bodyNumber = parsedData[i].bodyNumber;
        const chartTitle = `پروفیل بار مشترک: ${customerName} (${bodyNumber})`;

        try {
            // Render the whole chart-wrapper div to include button etc., or just canvas if preferred
            const canvas = await html2canvas(chartWrapper, { // Use chartWrapper for better context
                scale: 2, // Increase scale for better resolution
                useCORS: true // Important for images loaded from different origins if any
            });

            const imgData = canvas.toDataURL('image/png');
            const imgHeightCalculated = (canvas.height * chartRenderWidth) / canvas.width;

            // Check if there's enough space on the current page
            if (yOffset + imgHeightCalculated + 10 > pageHeight - margin) { // 10 for padding
                doc.addPage();
                yOffset = margin;
            }

            // Add chart image
            doc.addImage(imgData, 'PNG', margin, yOffset, chartRenderWidth, imgHeightCalculated);
            yOffset += imgHeightCalculated + 10; // Add padding after image

            showProgress(10 + ((i + 1) / charts.length) * 80); // Update progress
        } catch (error) {
            console.error(`Error generating PDF for chart ${i}:`, error);
            Swal.fire('خطا', `مشکلی در تولید PDF برای نمودار مشترک ${customerName} پیش آمد.`, 'error');
            Swal.close();
            showProgress(0);
            return;
        }
    }

    doc.save('پروفیل_بارهای_مشترکین.pdf');
    Swal.close();
    Swal.fire('موفقیت', 'فایل PDF با موفقیت دانلود شد.', 'success');
    showProgress(0); // Hide progress bar after completion
}

// Initialize application state on load
document.addEventListener('DOMContentLoaded', resetApplication);