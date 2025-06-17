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
        if (percent === 0 || percent === 100) {
            // progressContainer.style.display = 'none'; // Keep visible or hide based on preference
        } else {
            // progressContainer.style.display = 'block'; // Keep visible or hide based on preference
        }
    }
}

/**
 * تابع نابود کردن (حذف) تمام نمونه‌های نمودار Chart.js موجود
 * این کار برای جلوگیری از انباشت نمودارها در حافظه و مشکلات عملکردی ضروری است.
 */
function destroyCharts() {
    currentCharts.forEach(chart => {
        if (chart) {
            chart.destroy();
        }
    });
    currentCharts = []; // آرایه نمودارها را خالی می‌کند
    const chartsGrid = document.querySelector('.charts-grid'); // Changed to charts-grid
    if (chartsGrid) {
        chartsGrid.innerHTML = ''; // پاک کردن تمام نمودارها از DOM
    }
}

// ====================================================================================================
// توابع اصلی برنامه
// این توابع جریان اصلی برنامه را مدیریت می‌کنند و با رویدادهای کاربری تعامل دارند.
// ====================================================================================================

/**
 * تابع بارگذاری فایل اکسل
 * فایل انتخاب شده را می‌خواند و شیت‌های آن را برای انتخاب دراپ‌داون بارگذاری می‌کند.
 */
async function loadExcelFile() {
    showProgress(0, 'در حال بارگذاری فایل...');
    const file = fileInput.files[0];
    if (!file) {
        Swal.fire('خطا', 'لطفاً یک فایل اکسل انتخاب کنید!', 'error');
        showProgress(0, 'فایل انتخاب نشده...');
        return;
    }

    fileNameDisplay.textContent = file.name;
    sheetSelect.innerHTML = ''; // Clear previous options
    sheetSelect.disabled = true;
    processBtn.disabled = true;

    const reader = new FileReader();
    reader.onload = async (e) => {
        showProgress(20, 'در حال پردازش فایل...');
        const data = new Uint8Array(e.target.result);
        try {
            workbook = XLSX.read(data, { type: 'array' });
            workbook.SheetNames.forEach(sheetName => {
                const option = document.createElement('option');
                option.value = sheetName;
                option.textContent = sheetName;
                sheetSelect.appendChild(option);
            });
            sheetSelect.disabled = false;
            processBtn.disabled = false;
            showProgress(100, 'فایل آماده پردازش است.');
            Swal.fire('موفقیت', 'فایل اکسل با موفقیت بارگذاری شد.', 'success');
        } catch (error) {
            console.error('Error reading Excel file:', error);
            Swal.fire('خطا', 'فایل اکسل نامعتبر است یا در خواندن آن خطایی رخ داد.', 'error');
            showProgress(0, 'خطا در بارگذاری فایل...');
            fileNameDisplay.textContent = 'فایل انتخاب نشده...';
            workbook = null;
        }
    };
    reader.onerror = (e) => {
        console.error('FileReader error:', e);
        Swal.fire('خطا', 'خطا در خواندن فایل رخ داد.', 'error');
        showProgress(0, 'خطا در بارگذاری فایل...');
        fileNameDisplay.textContent = 'فایل انتخاب نشده...';
        workbook = null;
    };
    reader.readAsArrayBuffer(file);
}

/**
 * تابع اصلی برای پردازش داده‌ها از شیت انتخاب شده و به‌روزرسانی جدول و نمودارها
 */
async function processData() {
    if (!workbook) {
        Swal.fire('خطا', 'لطفاً ابتدا فایل اکسل را بارگذاری کنید.', 'error');
        return;
    }

    const selectedSheetName = sheetSelect.value;
    if (!selectedSheetName) {
        Swal.fire('خطا', 'لطفاً یک شیت را انتخاب کنید.', 'error');
        return;
    }

    showProgress(0, 'در حال استخراج داده‌ها...');
    const worksheet = workbook.Sheets[selectedSheetName];
    // Convert sheet to array of arrays (data rows)
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    if (jsonData.length < 2) {
        Swal.fire('خطا', 'شیت انتخاب شده حاوی داده‌ای نیست.', 'error');
        showProgress(0, 'پردازش متوقف شد.');
        return;
    }

    // Assume the first row is headers
    const headers = jsonData[0];
    const dataRows = jsonData.slice(1);

    parsedData = [];
    resultsTableBody.innerHTML = ''; // Clear existing table rows
    destroyCharts(); // Clear existing charts
    noChartsMessage.style.display = 'none'; // Hide no charts message

    const customerNameColIndex = headers.indexOf('نام مشترک');
    const customerIdColIndex = headers.indexOf('شناسه');
    const loadDataStartIndex = headers.indexOf('H1'); // Assuming load data starts from H1

    if (customerNameColIndex === -1 || customerIdColIndex === -1 || loadDataStartIndex === -1) {
        Swal.fire('خطا', 'فایل اکسل فرمت معتبری ندارد. ستون‌های "نام مشترک"، "شناسه" و "H1" یافت نشدند.', 'error');
        showProgress(0, 'پردازش متوقف شد.');
        return;
    }

    showProgress(20, 'در حال محاسبه و فیلتر...');

    const hourSelect = document.getElementById('hourSelect');
    const minuteSelect = document.getElementById('minuteSelect');
    const chkEvening = document.getElementById('chkEvening');
    const txtEvening = document.getElementById('txtEvening');
    const chkReduction = document.getElementById('chkReduction');
    const txtReduction = document.getElementById('txtReduction');

    const eveningStartHour = parseInt(hourSelect.value);
    const eveningStartMinute = parseInt(minuteSelect.value);
    const eveningLoadThreshold = chkEvening.checked ? parseFloat(txtEvening.value) : null;
    const maxReductionPercentage = chkReduction.checked ? parseFloat(txtReduction.value) : null; // Changed to max reduction

    const totalCustomers = dataRows.length;
    let processedCount = 0;

    for (const row of dataRows) {
        processedCount++;
        const name = row[customerNameColIndex] || 'نامشخص';
        const id = row[customerIdColIndex] || 'نامشخص';
        const loadData = row.slice(loadDataStartIndex, loadDataStartIndex + 24).map(val => parseFloat(val) || 0);

        // Calculate Morning Load (H1 to H12)
        const morningLoad = loadData.slice(0, 12).reduce((sum, val) => sum + val, 0);

        // Calculate Evening Load (from eveningStartHour to H24)
        let eveningLoad = 0;
        const eveningStartIndex = eveningStartHour; // H1 is index 0, H12 is index 11 etc.

        // Adjust index for evening load calculation if needed based on minuteSelect
        // For simplicity, assuming full hour. If minute matters, more complex logic is needed.
        if (eveningStartIndex < 24) { // Ensure index is within bounds
            eveningLoad = loadData.slice(eveningStartIndex, 24).reduce((sum, val) => sum + val, 0);
        }

        const reductionAmount = morningLoad - eveningLoad;
        const reductionPercentage = morningLoad > 0 ? (reductionAmount / morningLoad) * 100 : 0;

        // Filtering Logic (changed for max reduction)
        if (eveningLoadThreshold !== null && eveningLoad < eveningLoadThreshold) {
            continue; // Skip if evening load is below threshold
        }
        if (maxReductionPercentage !== null && reductionPercentage > maxReductionPercentage) {
            continue; // Skip if reduction percentage is ABOVE max allowed
        }

        parsedData.push({
            name,
            id,
            morningLoad,
            eveningLoad,
            reductionAmount,
            reductionPercentage,
            loadData // Keep original 24-hour data for charts
        });

        // Add to table
        const rowElement = resultsTableBody.insertRow();
        rowElement.insertCell().textContent = processedCount; // Row number
        rowElement.insertCell().textContent = name;
        rowElement.insertCell().textContent = id;
        rowElement.insertCell().textContent = morningLoad.toFixed(2);
        rowElement.insertCell().textContent = eveningLoad.toFixed(2);
        rowElement.insertCell().textContent = reductionAmount.toFixed(2);
        rowElement.insertCell().textContent = reductionPercentage.toFixed(2) + '%';

        // Update progress bar
        showProgress(20 + (processedCount / totalCustomers) * 70, `پردازش مشترک ${processedCount} از ${totalCustomers}`);
    }

    if (parsedData.length === 0) {
        Swal.fire('اطلاعات', 'هیچ مشترکی با فیلترهای اعمال شده یافت نشد.', 'info');
        noChartsMessage.style.display = 'block';
    } else {
        showProgress(95, 'در حال ترسیم نمودارها...');
        const chartAddress = document.getElementById('chartAddress').value;
        const addressColor = document.getElementById('addressColor').value;
        parsedData.forEach(customer => {
            createChart(customer, chartAddress, addressColor);
        });
        Swal.fire('موفقیت', 'داده‌ها با موفقیت پردازش و نمودارها ایجاد شدند.', 'success');
    }

    exportExcelBtn.disabled = parsedData.length === 0;
    exportPdfBtn.disabled = parsedData.length === 0;
    showProgress(100, 'پردازش کامل شد.');
}

/**
 * تابع ایجاد و نمایش نمودار برای یک مشترک خاص
 * @param {object} customerData - داده‌های مصرفی مشترک
 * @param {string} addressText - متنی برای نمایش به عنوان آدرس روی نمودار
 * @param {string} addressColor - رنگ متن آدرس
 */
function createChart(customerData, addressText = '', addressColor = '#FF0000') {
    const chartsGrid = document.querySelector('.charts-grid'); // Target the new grid container
    if (!chartsGrid) return;

    const chartWrapper = document.createElement('div');
    chartWrapper.className = 'chart-wrapper';

    const canvas = document.createElement('canvas');
    chartWrapper.appendChild(canvas);
    chartsGrid.appendChild(chartWrapper);

    const ctx = canvas.getContext('2d');

    // Generate labels for H1 to H24
    const labels = Array.from({ length: 24 }, (_, i) => `H${i + 1}`);

    const data = {
        labels: labels,
        datasets: [{
            label: `مصرف KW مشترک ${customerData.name} (${customerData.id})`,
            data: customerData.loadData,
            borderColor: 'rgba(75, 192, 192, 1)',
            backgroundColor: 'rgba(75, 192, 192, 0.2)', // Light fill
            pointRadius: 3, // Show points
            pointBackgroundColor: 'rgba(75, 192, 192, 1)',
            pointBorderColor: '#fff',
            pointHoverRadius: 5,
            borderWidth: 2,
            tension: 0.3, // Curve the line for a nicer look
            fill: true, // Fill area under the line
        }]
    };

    const config = {
        type: 'line',
        data: data,
        options: {
            responsive: true,
            maintainAspectRatio: false, // Important for controlling height via CSS
            plugins: {
                legend: {
                    position: 'top',
                    labels: {
                        font: {
                            family: 'Vazirmatn'
                        }
                    }
                },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                    rtl: true, // Enable RTL for tooltips
                    titleFont: { family: 'Vazirmatn' },
                    bodyFont: { family: 'Vazirmatn' },
                    callbacks: {
                        title: function(context) {
                            return `ساعت: ${context[0].label}`;
                        },
                        label: function(context) {
                            let label = context.dataset.label || '';
                            if (label) {
                                label += ': ';
                            }
                            if (context.parsed.y !== null) {
                                label += new Intl.NumberFormat('fa-IR').format(context.parsed.y) + ' KW';
                            }
                            return label;
                        }
                    }
                },
                annotation: {
                    annotations: {
                        addressLabel: {
                            type: 'label',
                            xValue: labels[Math.floor(labels.length / 2)], // Center horizontally (e.g., H12)
                            yValue: (ctx) => {
                                const chart = ctx.chart;
                                const scale = chart.scales.y;
                                // Position slightly below the top of the y-axis
                                return scale.max - (scale.max - scale.min) * 0.05;
                            },
                            content: addressText,
                            font: {
                                size: 10, // Significantly smaller font for address
                                weight: 'normal',
                                family: 'Vazirmatn'
                            },
                            color: addressColor, // Use selected color
                            backgroundColor: 'rgba(255, 255, 255, 0.6)', // Semi-transparent background
                            borderColor: addressColor,
                            borderWidth: 0.5, // Thinner border
                            borderRadius: 4, // Smaller border radius
                            xAdjust: 0,
                            yAdjust: 0,
                            display: addressText ? true : false, // Only display if text exists
                            position: 'start' // Align text to start
                        }
                    }
                }
            },
            scales: {
                x: {
                    title: {
                        display: true,
                        text: 'ساعت (H)',
                        font: {
                            family: 'Vazirmatn',
                            size: 12,
                            weight: 'bold'
                        }
                    },
                    grid: {
                        color: 'rgba(0, 0, 0, 0.05)', // Lighter grid lines
                    },
                    ticks: {
                        font: {
                            family: 'Vazirmatn',
                            size: 10
                        }
                    }
                },
                y: {
                    title: {
                        display: true,
                        text: 'میزان مصرف (KW)',
                        font: {
                            family: 'Vazirmatn',
                            size: 12,
                            weight: 'bold'
                        }
                    },
                    beginAtZero: true,
                    grid: {
                        color: 'rgba(0, 0, 0, 0.05)', // Lighter grid lines
                    },
                    ticks: {
                        callback: function(value) {
                            return new Intl.NumberFormat('fa-IR').format(value); // Format Y-axis numbers
                        },
                        font: {
                            family: 'Vazirmatn',
                            size: 10
                        }
                    }
                }
            },
            hover: {
                mode: 'nearest',
                intersect: true
            }
        },
    };

    const chart = new Chart(ctx, config);
    currentCharts.push(chart);
}

/**
 * تابع خروجی گرفتن اطلاعات جدول به اکسل
 */
async function exportToExcel() {
    if (parsedData.length === 0) {
        Swal.fire('خطا', 'هیچ داده‌ای برای خروجی گرفتن وجود ندارد.', 'error');
        return;
    }

    showProgress(0, 'در حال آماده‌سازی فایل اکسل...');
    const ws = XLSX.utils.json_to_sheet(parsedData.map(d => ({
        'نام مشترک': d.name,
        'شناسه': d.id,
        'بار صبح (KW)': d.morningLoad.toFixed(2),
        'بار عصر (KW)': d.eveningLoad.toFixed(2),
        'میزان کاهش (KW)': d.reductionAmount.toFixed(2),
        'درصد کاهش (%)': d.reductionPercentage.toFixed(2)
    })));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "نتایج تحلیل");

    // Add load data for each customer in separate sheets or append to main sheet
    parsedData.forEach(customer => {
        const loadHeaders = Array.from({ length: 24 }, (_, i) => `H${i + 1}`);
        const customerLoadData = [loadHeaders, customer.loadData];
        const customerWs = XLSX.utils.aoa_to_sheet(customerLoadData);
        XLSX.utils.book_append_sheet(wb, customerWs, `${customer.id}_بار`);
    });

    showProgress(70, 'در حال ذخیره فایل...');
    const fileName = `نتایج_تحلیل_مصرف_${new Date().toLocaleDateString('fa-IR')}.xlsx`;
    XLSX.writeFile(wb, fileName);
    showProgress(100, 'فایل اکسل با موفقیت ایجاد شد.');
    Swal.fire('موفقیت', 'فایل اکسل با موفقیت ایجاد و دانلود شد.', 'success');
}

/**
 * تابع خروجی گرفتن نمودارها به PDF
 */
async function exportChartsToPdf() {
    if (currentCharts.length === 0) {
        Swal.fire('خطا', 'هیچ نموداری برای خروجی گرفتن وجود ندارد.', 'error');
        return;
    }

    showProgress(0, 'در حال آماده‌سازی PDF...');

    // Load jsPDF dynamically (or include it in HTML if always needed)
    if (typeof jspdf === 'undefined') {
        Swal.fire({
            title: 'در حال بارگذاری کتابخانه PDF...',
            text: 'لطفاً صبر کنید.',
            allowOutsideClick: false,
            didOpen: () => {
                Swal.showLoading();
            }
        });
        const script = document.createElement('script');
        script.src = 'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
        document.head.appendChild(script);
        await new Promise(resolve => script.onload = resolve);
        Swal.close();
    }
    const { jsPDF } = window.jspdf;

    const doc = new jsPDF({
        orientation: 'p', // portrait
        unit: 'mm',
        format: 'a4'
    });

    const chartsToProcess = currentCharts.length;
    let processedCharts = 0;
    const pageHeight = doc.internal.pageSize.height;
    const pageWidth = doc.internal.pageSize.width;
    const margin = 10;
    let yPos = margin;
    let chartWidth = (pageWidth - 4 * margin) / 3; // 3 charts per row
    let chartHeight = 60; // Adjusted height for PDF (adjust as needed)
    let xOffset = margin;
    let chartCounter = 0;

    for (const chartInstance of currentCharts) {
        processedCharts++;
        showProgress( (processedCharts / chartsToProcess) * 100, `در حال افزودن نمودار ${processedCharts} از ${chartsToProcess} به PDF...`);

        const imgData = chartInstance.canvas.toDataURL('image/png', 1.0); // Get image data from canvas

        // Calculate position for 3 charts per row
        if (chartCounter > 0 && chartCounter % 3 === 0) {
            yPos += chartHeight + margin;
            xOffset = margin;
            if (yPos + chartHeight + margin > pageHeight) {
                doc.addPage();
                yPos = margin;
            }
        }
        
        doc.addImage(imgData, 'PNG', xOffset, yPos, chartWidth, chartHeight);
        xOffset += chartWidth + margin;
        chartCounter++;
    }

    doc.save(`پروفیل_بارهای_مشترکین_${new Date().toLocaleDateString('fa-IR')}.pdf`);
    showProgress(100, 'فایل PDF با موفقیت ایجاد شد.');
    Swal.fire('موفقیت', 'فایل PDF پروفیل‌ها با موفقیت ایجاد و دانلود شد.', 'success');
}


// ====================================================================================================
// رویدادها (Event Listeners)
// این بخش رویدادهای DOM را مدیریت می‌کند و توابع مربوطه را فراخوانی می‌کند.
// ====================================================================================================

// DOM Elements
const fileInput = document.getElementById('excelFile');
const fileNameDisplay = document.getElementById('fileNameDisplay');
const sheetSelect = document.getElementById('sheetSelect');
const processBtn = document.getElementById('processDataBtn');
const resetAppBtn = document.getElementById('resetAppBtn');
const resultsTableBody = document.querySelector('#resultsTable tbody');
const chartsContainer = document.getElementById('chartsContainer');
const noChartsMessage = document.getElementById('noChartsMessage');
const exportExcelBtn = document.getElementById('exportExcelBtn');
const exportPdfBtn = document.getElementById('exportPdfBtn');

const chkEvening = document.getElementById('chkEvening');
const txtEvening = document.getElementById('txtEvening');
const chkReduction = document.getElementById('chkReduction');
const txtReduction = document.getElementById('txtReduction');
const hourSelect = document.getElementById('hourSelect');
const minuteSelect = document.getElementById('minuteSelect');


// Initialize hour and minute selects
for (let i = 1; i <= 24; i++) {
    const option = document.createElement('option');
    option.value = i;
    option.textContent = i;
    hourSelect.appendChild(option);
}
for (let i = 0; i < 60; i += 15) { // Every 15 minutes
    const option = document.createElement('option');
    option.value = i;
    option.textContent = i < 10 ? `0${i}` : i;
    minuteSelect.appendChild(option);
}
// Set default values
hourSelect.value = 17; // Default 17 for evening
minuteSelect.value = 0;


// Event Listeners
if (fileInput) fileInput.addEventListener('change', loadExcelFile);
if (processBtn) processBtn.addEventListener('click', processData);
if (exportExcelBtn) exportExcelBtn.addEventListener('click', exportToExcel);
if (exportPdfBtn) exportPdfBtn.addEventListener('click', exportChartsToPdf);

// Control enable/disable of number inputs based on checkbox
if (chkEvening) chkEvening.addEventListener('change', () => {
    if (txtEvening) txtEvening.disabled = !chkEvening.checked;
});
if (chkReduction) chkReduction.addEventListener('change', () => {
    if (txtReduction) txtReduction.disabled = !chkReduction.checked;
});

if (resetAppBtn) {
    resetAppBtn.addEventListener('click', () => {
        workbook = null;
        parsedData = [];
        currentCharts = [];

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
        destroyCharts();
        if (noChartsMessage) noChartsMessage.style.display = 'block';

        if (chkEvening) chkEvening.checked = false;
        if (txtEvening) txtEvening.disabled = true;
        if (chkReduction) chkReduction.checked = false;
        if (txtReduction) txtReduction.disabled = true;

        // Reset address inputs
        const chartAddress = document.getElementById('chartAddress');
        const addressColor = document.getElementById('addressColor');
        if (chartAddress) chartAddress.value = '';
        if (addressColor) addressColor.value = '#FF0000';

        showProgress(0, 'منتظر انتخاب فایل...');
        Swal.fire('با موفقیت', 'برنامه به حالت اولیه بازنشانی شد.', 'success');
    });
}

// Initial state on load
document.addEventListener('DOMContentLoaded', () => {
    showProgress(0, 'برنامه آماده است. فایل اکسل را انتخاب کنید.');
    // Set initial state of checkboxes and inputs
    if (txtEvening) txtEvening.disabled = true;
    if (txtReduction) txtReduction.disabled = true;
});