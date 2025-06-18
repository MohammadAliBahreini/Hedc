// ====================================================================================================
// تعریف متغیرهای سراسری
// ====================================================================================================
let workbook;
let parsedData = [];
let currentCharts = [];
let filteredData = []; // برای نگهداری داده‌های فیلتر شده
let appLogs = []; // آرایه برای نگهداری لاگ‌های برنامه

// تعریف ستون‌های مورد انتظار دقیقا بر اساس لیست شما
const REQUIRED_COLUMNS = [
    '#', 'Serial no.', 'Date', 'Customer name', 'Billing id', 'Customer id', 'Contracted demand', 'Address',
    '00:00 to 00:15 [KW]', '00:15 to 00:30 [KW]', '00:30 to 00:45 [KW]', '00:45 to 01:00 [KW]',
    '01:00 to 01:15 [KW]', '01:15 to 01:30 [KW]', '01:30 to 01:45 [KW]', '01:45 to 02:00 [KW]',
    '02:00 to 02:15 [KW]', '02:15 to 02:30 [KW]', '02:30 to 02:45 [KW]', '02:45 to 03:00 [KW]',
    '03:00 to 03:15 [KW]', '03:15 to 03:30 [KW]', '03:30 to 03:45 [KW]', '03:45 to 04:00 [KW]',
    '04:00 to 04:15 [KW]', '04:15 to 04:30 [KW]', '04:30 to 04:45 [KW]', '04:45 to 05:00 [KW]',
    '05:00 to 05:15 [KW]', '05:15 to 05:30 [KW]', '05:30 to 05:45 [KW]', '05:45 to 06:00 [KW]',
    '06:00 to 06:15 [KW]', '06:15 to 06:30 [KW]', '06:30 to 06:45 [KW]', '06:45 to 07:00 [KW]',
    '07:00 to 07:15 [KW]', '07:15 to 07:30 [KW]', '07:30 to 07:45 [KW]', '07:45 to 08:00 [KW]',
    '08:00 to 08:15 [KW]', '08:15 to 08:30 [KW]', '08:30 to 08:45 [KW]', '08:45 to 09:00 [KW]',
    '09:00 to 09:15 [KW]', '09:15 to 09:30 [KW]', '09:30 to 09:45 [KW]', '09:45 to 10:00 [KW]',
    '10:00 to 10:15 [KW]', '10:15 to 10:30 [KW]', '10:30 to 10:45 [KW]', '10:45 to 11:00 [KW]',
    '11:00 to 11:15 [KW]', '11:15 to 11:30 [KW]', '11:30 to 11:45 [KW]', '11:45 to 12:00 [KW]',
    '12:00 to 12:15 [KW]', '12:15 to 12:30 [KW]', '12:30 to 12:45 [KW]', '12:45 to 13:00 [KW]',
    '13:00 to 13:15 [KW]', '13:15 to 13:30 [KW]', '13:30 to 13:45 [KW]', '13:45 to 14:00 [KW]',
    '14:00 to 14:15 [KW]', '14:15 to 14:30 [KW]', '14:30 to 14:45 [KW]', '14:45 to 15:00 [KW]',
    '15:00 to 15:15 [KW]', '15:15 to 15:30 [KW]', '15:30 to 15:45 [KW]', '15:45 to 16:00 [KW]',
    '16:00 to 16:15 [KW]', '16:15 to 16:30 [KW]', '16:30 to 16:45 [KW]', '16:45 to 17:00 [KW]',
    '17:00 to 17:15 [KW]', '17:15 to 17:30 [KW]', '17:30 to 17:45 [KW]', '17:45 to 18:00 [KW]',
    '18:00 to 18:15 [KW]', '18:15 to 18:30 [KW]', '18:30 to 18:45 [KW]', '18:45 to 19:00 [KW]',
    '19:00 to 19:15 [KW]', '19:15 to 19:30 [KW]', '19:30 to 19:45 [KW]', '19:45 to 20:00 [KW]',
    '20:00 to 20:15 [KW]', '20:15 to 20:30 [KW]', '20:30 to 20:45 [KW]', '20:45 to 21:00 [KW]',
    '21:00 to 21:15 [KW]', '21:15 to 21:30 [KW]', '21:30 to 21:45 [KW]', '21:45 to 22:00 [KW]',
    '22:00 to 22:15 [KW]', '22:15 to 22:30 [KW]', '22:30 to 22:45 [KW]', '22:45 to 23:00 [KW]',
    '23:00 to 23:15 [KW]', '23:15 to 23:30 [KW]', '23:30 to 23:45 [KW]', '23:45 to 00:00 [KW]'
];

// ستون‌هایی که باید در جدول نمایش داده شوند (مصرف‌های ۱۵ دقیقه‌ای حذف شدند)
const DISPLAY_COLUMNS = [
    '#', 'Serial no.', 'Date', 'Customer name', 'Billing id', 'Customer id', 'Contracted demand', 'Address',
    'Morning Load (KW)', 'Evening Load (KW)', 'Reduction Amount (KW)', 'Reduction Percentage (%)'
    // 'Total Consumption [KWh]', 'Average consumption [KW]', 'Max consumption [KW]', 'Min consumption [KW]',
    // 'Consumption per contracted demand (%)', 'Consumption per average consumption (%)', 'Consumption per max consumption (%)',
    // 'Load Factor (LF) [%]', 'Diversity Factor (DF) [%]', 'Coincidence Factor (CF) [%]',
    // 'Demand Factor (DMF) [%]', 'Peak Hour'
];

// تعریف عناصر صفحه به صورت مرکزی
const elements = {
    processBtn: document.getElementById('processDataBtn'),
    resetBtn: document.getElementById('resetAppBtn'),
    excelFile: document.getElementById('excelFile'),
    sheetSelect: document.getElementById('sheetSelect'),
    dataTableBody: document.getElementById('dataTableBody'),
    fileNameDisplay: document.getElementById('fileNameDisplay'),
    chartsContainer: document.getElementById('chartsContainer'),
    noChartsMessage: document.getElementById('noChartsMessage'),
    exportPdfBtn: document.getElementById('exportPdfBtn'),
    exportExcelBtn: document.getElementById('exportExcelBtn'),
    exportChartsAsImagesBtn: document.getElementById('exportChartsAsImagesBtn'),
    exportLogFileBtn: document.getElementById('exportLogFile'),
    searchInput: document.getElementById('filterValue'), // توجه: در کد شما این شناسه تغییر کرده
    applyFilterBtn: document.getElementById('applyFilterBtn'),
    clearFilterBtn: document.getElementById('clearFilterBtn'),
    minConsumptionInput: document.getElementById('txtEvening'),
    maxConsumptionInput: document.getElementById('txtReduction'),
    // filterConsumptionBtn: document.getElementById('filterConsumptionBtn'),
    // clearConsumptionFilterBtn: document.getElementById('clearConsumptionFilterBtn'),
    timePeriodSelect: document.getElementById('timePeriodSelect'),
    calculateTimePeriodBtn: document.getElementById('calculateTimePeriodBtn'),
    timePeriodResultDiv: document.getElementById('timePeriodResult'),
    renderAllChartsBtn: document.getElementById('renderAllChartsBtn'),
    dataTable: document.getElementById('resultsTable')
};

// بررسی وجود همه عناصر
function validateElements() {
    let allElementsExist = true;
    for (const [name, element] of Object.entries(elements)) {
        if (!element) {
            console.error(`عنصر ${name} یافت نشد!`);
            allElementsExist = false;
            
            // نمایش پیام خطا فقط در حالت توسعه
            // if (process.env.NODE_ENV === 'development') {
                // alert(`عنصر ${name} در صفحه یافت نشد!`);
            // }
        }
    }
    return allElementsExist;
}

// اجرای بررسی عناصر هنگام بارگذاری صفحه
document.addEventListener('DOMContentLoaded', function() {
    if (!validateElements()) {
        console.error('برخی عناصر ضروری در صفحه یافت نشدند!');
        // در اینجا می‌توانید تصمیم بگیرید که چگونه با خطا برخورد کنید
    }
});
/**
 * تابع لاگ برای ثبت رویدادها
 * @param {string} level - سطح لاگ (e.g., 'info', 'warn', 'error')
 * @param {string} message - پیام لاگ
 */
function log(level, message) {
    const timestamp = new Date().toLocaleString('fa-IR', { timeZone: 'Asia/Tehran' });
    appLogs.push({ timestamp, level, message });
    console.log(`[${level.toUpperCase()}] ${timestamp}: ${message}`);
}

/**
 * نمایش پیام در Swal
 * @param {string} title
 * @param {string} text
 * @param {string} icon
 */
function showAlert(title, text, icon) {
    Swal.fire({
        title: title,
        text: text,
        icon: icon,
        confirmButtonText: 'باشه'
    });
}

/**
 * اعتبارسنجی ستون‌ها
 * @param {Array<string>} headers - هدرهای فایل اکسل
 * @returns {boolean} - آیا هدرها معتبر هستند یا خیر
 */
function validateHeaders(headers) {
    const missingColumns = REQUIRED_COLUMNS.filter(col => !headers.includes(col));
    if (missingColumns.length > 0) {
        log('error', `ستون‌های زیر در فایل اکسل یافت نشدند: ${missingColumns.join(', ')}`);
        showAlert('خطا در ساختار فایل', `ستون‌های زیر در فایل اکسل یافت نشدند:<br>${missingColumns.join(', ')}<br>لطفاً فایل صحیح را بارگذاری کنید.`, 'error');
        return false;
    }
    log('info', 'ساختار ستون‌های فایل اکسل معتبر است.');
    return true;
}

/**
 * پردازش داده‌های اکسل
 * @param {Array<Object>} data - داده‌های خوانده شده از اکسل
 * @returns {Array<Object>} - داده‌های پردازش شده
 */
// اصلاح تابع processExcelData برای محاسبه مقادیر
function processExcelData(data) {
    const processedData = [];
    
    // دریافت تنظیمات زمان از UI
    const morningStart = document.getElementById('morningPeakStart').value;
    const morningEnd = document.getElementById('morningPeakEnd').value;
    const eveningStart = document.getElementById('eveningPeakStart').value;
    const eveningEnd = document.getElementById('eveningPeakEnd').value;
    
    data.forEach((row, index) => {
        // محاسبه بار صبح
        const morningLoad = calculatePeakLoad(row, morningStart, morningEnd, 'morningCalcType');
        // محاسبه بار عصر
        const eveningLoad = calculatePeakLoad(row, eveningStart, eveningEnd, 'eveningCalcType');
        
        const reductionKW = morningLoad - eveningLoad;
        const reductionPercent = morningLoad > 0 ? (reductionKW / morningLoad) * 100 : 0;
        
        const processedRow = {
            // ... سایر فیلدها
            'Morning Load (KW)': morningLoad.toFixed(2),
            'Evening Load (KW)': eveningLoad.toFixed(2),
            'Reduction Amount (KW)': reductionKW.toFixed(2),
            'Reduction Percentage (%)': reductionPercent.toFixed(2),
            // ... سایر فیلدها
        };
        
        processedData.push(processedRow);
    });
    
    return processedData;
}

function calculatePeakLoad(row, startTime, endTime, calcTypeId) {
    const calcType = document.getElementById(calcTypeId).value;
    const loads = [];
    
    // تبدیل زمان به دقیقه
    const toMinutes = (timeStr) => {
        const [h, m] = timeStr.split(':').map(Number);
        return h * 60 + m;
    };
    
    const startMin = toMinutes(startTime);
    const endMin = toMinutes(endTime);
    
    for (let i = 0; i < 24; i++) {
        for (let j = 0; j < 4; j++) {
            const currentMin = i * 60 + j * 15;
            if (currentMin >= startMin && currentMin <= endMin) {
                const hour = String(i).padStart(2, '0');
                const minute = String(j * 15).padStart(2, '0');
                const nextMin = String((j + 1) * 15).padStart(2, '0');
                const colName = `${hour}:${minute} to ${hour}:${nextMin} [KW]`;
                const value = parseFloat(row[colName]) || 0;
                loads.push(value);
            }
        }
    }
    
    if (loads.length === 0) return 0;
    
    switch (calcType) {
        case 'avg': return loads.reduce((a, b) => a + b, 0) / loads.length;
        case 'max': return Math.max(...loads);
        case 'min': return Math.min(...loads);
        default: return loads.reduce((a, b) => a + b, 0) / loads.length;
    }
}
function processExcelData2(data) {
    const processedData = [];
    data.forEach((row, index) => {
        // ... کدهای قبلی
        
        // محاسبه بار صبح و عصر
        const morningPeakStart = document.getElementById('morningPeakStart').value;
        const morningPeakEnd = document.getElementById('morningPeakEnd').value;
        const eveningPeakStart = document.getElementById('eveningPeakStart').value;
        const eveningPeakEnd = document.getElementById('eveningPeakEnd').value;
        
        // تبدیل زمان به دقیقه
        const toMinutes = (timeStr) => {
            const [h, m] = timeStr.split(':').map(Number);
            return h * 60 + m;
        };
        
        // محاسبه بار صبح
        const morningLoads = [];
        // محاسبه بار عصر
        const eveningLoads = [];
        
        for (let i = 0; i < 24; i++) {
            for (let j = 0; j < 4; j++) {
                const hour = String(i).padStart(2, '0');
                const minute = String(j * 15).padStart(2, '0');
                const nextMinute = String((j + 1) * 15).padStart(2, '0');
                const colName = `${hour}:${minute} to ${hour}:${nextMinute} [KW]`;
                const currentTime = i * 60 + j * 15;
                
                if (currentTime >= toMinutes(morningPeakStart) && currentTime <= toMinutes(morningPeakEnd)) {
                    morningLoads.push(parseFloat(row[colName]) || 0);
                }
                
                if (currentTime >= toMinutes(eveningPeakStart) && currentTime <= toMinutes(eveningPeakEnd)) {
                    eveningLoads.push(parseFloat(row[colName]) || 0);
                }
            }
        }
        
        // محاسبه میانگین/حداکثر/حداقل بر اساس انتخاب کاربر
        const morningCalcType = document.getElementById('morningCalcType').value;
        const eveningCalcType = document.getElementById('eveningCalcType').value;
        
        const calculateLoad = (loads, calcType) => {
            if (loads.length === 0) return 0;
            switch (calcType) {
                case 'avg': return loads.reduce((a, b) => a + b, 0) / loads.length;
                case 'max': return Math.max(...loads);
                case 'min': return Math.min(...loads);
                default: return loads.reduce((a, b) => a + b, 0) / loads.length;
            }
        };
        
        const morningLoad = calculateLoad(morningLoads, morningCalcType);
        const eveningLoad = calculateLoad(eveningLoads, eveningCalcType);
        const reductionAmount = morningLoad - eveningLoad;
        const reductionPercent = (morningLoad > 0) ? (reductionAmount / morningLoad) * 100 : 0;
        
        const processedRow = {
            // ... سایر فیلدها
            'Morning Load (KW)': morningLoad.toFixed(2),
            'Evening Load (KW)': eveningLoad.toFixed(2),
            'Reduction Amount (KW)': reductionAmount.toFixed(2),
            'Reduction Percentage (%)': reductionPercent.toFixed(2),
            // ... سایر فیلدها
        };
        
        processedData.push(processedRow);
    });
    return processedData;
}
function drawCharts(dataToChart) {
    // پاک کردن نمودارهای قبلی
    currentCharts.forEach(chart => chart.destroy());
    currentCharts = [];
    
    const chartsGrid = document.querySelector('.charts-grid');
    if (chartsGrid) {
        chartsGrid.innerHTML = '';
    } else {
        console.error('عنصر charts-grid یافت نشد!');
        return;
    }

    if (dataToChart.length === 0) {
        if (noChartsMessage) {
            noChartsMessage.style.display = 'block';
        }
        log('warn', 'داده‌ای برای رسم نمودار وجود ندارد.');
        return;
    }
    
    if (noChartsMessage) {
        noChartsMessage.style.display = 'none';
    }

    // بقیه کد تابع بدون تغییر
    // ...
}
function drawCharts1(dataToChart) {
    // پاک کردن نمودارهای قبلی
    currentCharts.forEach(chart => chart.destroy());
    currentCharts = [];
    
    const chartsGrid = document.querySelector('.charts-grid');
    if (chartsGrid) {
        chartsGrid.innerHTML = '';
    } else {
        console.error('عنصر charts-grid یافت نشد!');
        return;
    }

    if (dataToChart.length === 0) {
        if (noChartsMessage) {
            noChartsMessage.style.display = 'block';
        }
        log('warn', 'داده‌ای برای رسم نمودار وجود ندارد.');
        return;
    }
    
    if (noChartsMessage) {
        noChartsMessage.style.display = 'none';
    }

    // بقیه کد تابع بدون تغییر
    // ...
}

function processExcelData1(data) {
    const processedData = [];
    data.forEach((row, index) => {
        // اطمینان از وجود ستون‌های لازم و تبدیل به نوع صحیح
        const serialNo = row['Serial no.'];
        const customerName = row['Customer name'];
        const billingId = row['Billing id'];
        const customerId = row['Customer id'];
        const contractedDemand = parseFloat(row['Contracted demand']);
        const address = row['Address'];
        const date = row['Date'];

        // جمع‌آوری داده‌های مصرف ۱۵ دقیقه‌ای
        let consumptionData = {};
        let totalConsumptionKWh = 0;
        let consumptionValues = []; // برای محاسبات میانگین، حداقل و حداکثر
        for (let i = 0; i < 24; i++) {
            for (let j = 0; j < 4; j++) {
                const hour = String(i).padStart(2, '0');
                const minute = String(j * 15).padStart(2, '0');
                const colName = `${hour}:${minute} to ${hour}:${String(parseInt(minute) + 15).padStart(2, '0')} [KW]`;
                const value = parseFloat(row[colName]);
                consumptionData[colName] = isNaN(value) ? 0 : value;
                // هر ۱۵ دقیقه یک چهارم ساعت است، پس برای تبدیل به KWh تقسیم بر 4 می‌کنیم.
                if (!isNaN(value)) {
                    totalConsumptionKWh += value / 4;
                    consumptionValues.push(value);
                }
            }
        }

        // محاسبه میانگین، حداکثر و حداقل مصرف
        const averageConsumptionKW = consumptionValues.length > 0 ? consumptionValues.reduce((a, b) => a + b, 0) / consumptionValues.length : 0;
        const maxConsumptionKW = consumptionValues.length > 0 ? Math.max(...consumptionValues) : 0;
        const minConsumptionKW = consumptionValues.length > 0 ? Math.min(...consumptionValues) : 0;

        // محاسبه درصد کاهش بر اساس توان قراردادی
        const consumptionPerContractedDemand = contractedDemand > 0 ? (totalConsumptionKWh / (contractedDemand * 24)) * 100 : 0; // مصرف در روز تقسیم بر توان قراردادی در 24 ساعت
        const consumptionPerAverageConsumption = averageConsumptionKW > 0 ? (totalConsumptionKWh / (averageConsumptionKW * 24)) * 100 : 0; // مصرف در روز تقسیم بر میانگین مصرف در 24 ساعت
        const consumptionPerMaxConsumption = maxConsumptionKW > 0 ? (totalConsumptionKWh / (maxConsumptionKW * 24)) * 100 : 0; // مصرف در روز تقسیم بر حداکثر مصرف در 24 ساعت

        // محاسبه Load Factor (LF)
        const loadFactor = (averageConsumptionKW > 0 && maxConsumptionKW > 0) ? (averageConsumptionKW / maxConsumptionKW) * 100 : 0;

        // محاسبه Diversity Factor (DF) - نیاز به داده‌های چندین مشترک است.
        // فعلا برای یک مشترک، برابر با ۱۰۰ درصد در نظر گرفته می‌شود یا نیاز به منطق پیچیده‌تری دارد.
        // برای سادگی، اگر برای یک مشترک محاسبه می‌شود، ممکن است از ۱۰۰% شروع کنیم.
        const diversityFactor = 100; // placeholder, needs multiple customer data

        // محاسبه Coincidence Factor (CF) - نیاز به داده‌های چندین مشترک است.
        // فعلا برای یک مشترک، برابر با ۱۰۰ درصد در نظر گرفته می‌شود یا نیاز به منطق پیچیده‌تری دارد.
        const coincidenceFactor = 100; // placeholder, needs multiple customer data

        // محاسبه Demand Factor (DMF)
        const demandFactor = (contractedDemand > 0) ? (maxConsumptionKW / contractedDemand) * 100 : 0;

        // پیدا کردن Peak Hour
        let peakHour = 'N/A';
        if (consumptionValues.length > 0) {
            const maxVal = Math.max(...consumptionValues);
            const maxIndex = consumptionValues.indexOf(maxVal);
            const peakHourStart = Math.floor(maxIndex / 4);
            const peakMinuteStart = (maxIndex % 4) * 15;
            const peakHourEnd = peakHourStart;
            const peakMinuteEnd = peakMinuteStart + 15;
            peakHour = `${String(peakHourStart).padStart(2, '0')}:${String(peakMinuteStart).padStart(2, '0')} تا ${String(peakHourEnd).padStart(2, '0')}:${String(peakMinuteEnd).padStart(2, '0')}`;
        }

        const processedRow = {
            '#': index + 1, // شماره ردیف
            'Serial no.': serialNo,
            'Date': date,
            'Customer name': customerName,
            'Billing id': billingId,
            'Customer id': customerId,
            'Contracted demand': contractedDemand,
            'Address': address,
        };
        processedData.push(processedRow);
    });
    log('info', `پردازش ${processedData.length} ردیف داده انجام شد.`);
    return processedData;
}


/**
 * رندر کردن داده‌ها در جدول
 * @param {Array<Object>} data - داده‌هایی که باید در جدول نمایش داده شوند
 */
function renderTable(data) {
    if (!dataTableBody) {
        log('error', 'عنصر tbody برای جدول یافت نشد.');
        return;
    }
    dataTableBody.innerHTML = ''; // پاک کردن محتویات قبلی جدول
    if (data.length === 0) {
        dataTableBody.innerHTML = '<tr><td colspan="15" style="text-align: center;">داده‌ای برای نمایش وجود ندارد.</td></tr>';
        log('info', 'جدول خالی رندر شد، داده‌ای برای نمایش نیست.');
        return;
    }

    data.forEach((row, index) => {
        const tr = document.createElement('tr');
        // نمایش ردیف‌های فرد و زوج با رنگ‌های متفاوت
        if (index % 2 === 0) {
            tr.classList.add('even-row');
        } else {
            tr.classList.add('odd-row');
        }

        // اضافه کردن ستون '#' (شماره ردیف)
        const rowNumTd = document.createElement('td');
        rowNumTd.textContent = index + 1; // شماره ردیف بر اساس ایندکس در آرایه
        tr.appendChild(rowNumTd);

        DISPLAY_COLUMNS.forEach(col => {
            if (col !== '#') { // شماره ردیف قبلاً اضافه شده
                const td = document.createElement('td');
                let cellValue = row[col];

                // فرمت کردن اعداد اعشاری
                if (typeof cellValue === 'number') {
                    cellValue = cellValue.toLocaleString('fa-IR', { maximumFractionDigits: 2 });
                }

                td.textContent = cellValue;
                tr.appendChild(td);
            }
        });

        // اضافه کردن دکمه "نمودار"
        const chartTd = document.createElement('td');
        const chartButton = document.createElement('button');
        chartButton.className = 'btn btn-info btn-sm';
        chartButton.innerHTML = '<i class="fas fa-chart-bar"></i>';
        chartButton.title = 'نمایش نمودار مصرف';
        chartButton.onclick = () => drawCharts([row]); // ارسال فقط ردیف فعلی برای رسم نمودار
        chartTd.appendChild(chartButton);
        tr.appendChild(chartTd);

        // اضافه کردن دکمه "حذف"
        const deleteTd = document.createElement('td');
        const deleteButton = document.createElement('button');
        deleteButton.className = 'btn btn-danger btn-sm';
        deleteButton.innerHTML = '<i class="fas fa-trash"></i>';
        deleteButton.title = 'حذف ردیف';
        deleteButton.onclick = () => deleteRow(row['Serial no.']); // حذف بر اساس شماره سریال
        deleteTd.appendChild(deleteButton);
        tr.appendChild(deleteTd);

        dataTableBody.appendChild(tr);
    });
    log('info', `نمایش ${data.length} ردیف در جدول.`);
}

/**
 * حذف یک ردیف از داده‌های فیلتر شده و بازخوانی جدول و نمودارها
 * @param {string} serialNo - شماره سریال ردیف مورد نظر برای حذف
 */
function deleteRow(serialNo) {
    Swal.fire({
        title: 'آیا مطمئن هستید؟',
        text: "این ردیف از جدول و نمودارها حذف خواهد شد!",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#3085d6',
        cancelButtonColor: '#d33',
        confirmButtonText: 'بله، حذف کن!',
        cancelButtonText: 'خیر'
    }).then((result) => {
        if (result.isConfirmed) {
            const initialLength = filteredData.length;
            filteredData = filteredData.filter(row => row['Serial no.'] !== serialNo);
            if (filteredData.length < initialLength) {
                renderTable(filteredData);
                drawCharts(filteredData); // بازسازی نمودارها با داده‌های به‌روز شده
                showAlert('حذف شد!', 'ردیف با موفقیت حذف شد.', 'success');
                log('info', `ردیف با شماره سریال ${serialNo} حذف شد.`);
            } else {
                showAlert('خطا', 'ردیف مورد نظر یافت نشد.', 'error');
                log('error', `تلاش برای حذف ردیف با شماره سریال ${serialNo} ناموفق بود (یافت نشد).`);
            }
        }
    });
}

/**
 * رسم نمودارها
 * @param {Array<Object>} dataToChart - داده‌هایی که باید برای رسم نمودار استفاده شوند
 */
function drawCharts(dataToChart) {
    // پاک کردن نمودارهای قبلی
    currentCharts.forEach(chart => chart.destroy());
    currentCharts = [];
    chartsContainer.innerHTML = ''; // پاک کردن محتویات قبلی

    if (dataToChart.length === 0) {
        noChartsMessage.style.display = 'block';
        log('warn', 'داده‌ای برای رسم نمودار وجود ندارد.');
        return;
    }
    noChartsMessage.style.display = 'none';

    dataToChart.forEach(rowData => {
        const customerName = rowData['Customer name'];
        const date = rowData['Date'];
        const contractedDemand = rowData['Contracted demand'];

        // استخراج داده‌های مصرف 15 دقیقه‌ای
        const labels = [];
        const consumptionValues = [];
        for (let i = 0; i < 24; i++) {
            for (let j = 0; j < 4; j++) {
                const hourStart = String(i).padStart(2, '0');
                const minuteStart = String(j * 15).padStart(2, '0');
                const hourEnd = String(i).padStart(2, '0');
                const minuteEnd = String((j + 1) * 15).padStart(2, '0');
                const colName = `${hourStart}:${minuteStart} to ${hourEnd}:${minuteEnd} [KW]`;
                labels.push(`${hourStart}:${minuteStart}`);
                consumptionValues.push(rowData[colName]);
            }
        }

        const chartId = `chart-${rowData['Serial no.']}`;
        const chartContainerDiv = document.createElement('div');
        chartContainerDiv.className = 'chart-container card';
        chartContainerDiv.innerHTML = `<h3>پروفیل بار ${customerName} در تاریخ ${date}</h3><canvas id="${chartId}"></canvas>`;
        chartsContainer.appendChild(chartContainerDiv);

        const ctx = document.getElementById(chartId).getContext('2d');
        const newChart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: labels,
                datasets: [{
                    label: 'مصرف [KW]',
                    data: consumptionValues,
                    borderColor: 'rgb(75, 192, 192)',
                    tension: 0.1,
                    fill: false
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    title: {
                        display: true,
                        text: `پروفیل بار مشترک: ${customerName} - تاریخ: ${date}`,
                        font: {
                            size: 16,
                            family: 'Vazirmatn'
                        }
                    },
                    tooltip: {
                        callbacks: {
                            title: function(context) {
                                return `ساعت: ${context[0].label}`;
                            },
                            label: function(context) {
                                return `مصرف: ${context.parsed.y} KW`;
                            }
                        },
                        bodyFont: {
                            family: 'Vazirmatn'
                        },
                        titleFont: {
                            family: 'Vazirmatn'
                        }
                    },
                    annotation: {
                        annotations: {
                            line1: {
                                type: 'line',
                                yMin: contractedDemand,
                                yMax: contractedDemand,
                                borderColor: 'rgb(255, 99, 132)',
                                borderWidth: 2,
                                borderDash: [5, 5],
                                label: {
                                    content: `توان قراردادی: ${contractedDemand} KW`,
                                    enabled: true,
                                    position: 'end',
                                    backgroundColor: 'rgba(255, 99, 132, 0.8)',
                                    font: {
                                        family: 'Vazirmatn'
                                    }
                                }
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        title: {
                            display: true,
                            text: 'زمان (ساعت)',
                            font: {
                                size: 14,
                                family: 'Vazirmatn'
                            }
                        },
                        ticks: {
                            font: {
                                family: 'Vazirmatn'
                            }
                        }
                    },
                    y: {
                        beginAtZero: true,
                        title: {
                            display: true,
                            text: 'مصرف (کیلووات)',
                            font: {
                                size: 14,
                                family: 'Vazirmatn'
                            }
                        },
                        ticks: {
                            font: {
                                family: 'Vazirmatn'
                            }
                        }
                    }
                }
            }
        });
        currentCharts.push(newChart);
    });
    log('info', `رسم ${dataToChart.length} نمودار انجام شد.`);
}

/**
 * فیلتر کردن داده‌ها بر اساس نام مشترک
 */
function filterDataByCustomerName() {
    const searchTerm = searchInput.value.toLowerCase();
    if (searchTerm === '') {
        filteredData = [...parsedData]; // اگر جستجو خالی است، همه داده‌های پردازش شده را نمایش بده
        log('info', 'فیلتر نام مشترک حذف شد.');
    } else {
        filteredData = parsedData.filter(row =>
            row['Customer name'] && row['Customer name'].toLowerCase().includes(searchTerm)
        );
        log('info', `فیلتر بر اساس نام مشترک: "${searchTerm}" اعمال شد. تعداد نتایج: ${filteredData.length}`);
    }
    renderTable(filteredData);
    drawCharts(filteredData);
}

/**
 * پاک کردن فیلتر نام مشترک
 */
function clearCustomerFilter() {
    searchInput.value = '';
    filterDataByCustomerName(); // بازنشانی فیلتر
    log('info', 'فیلتر نام مشترک پاک شد.');
}

/**
 * فیلتر کردن داده‌ها بر اساس محدوده مصرف کل
 */
function filterDataByConsumption() {
    const minVal = parseFloat(minConsumptionInput.value);
    const maxVal = parseFloat(maxConsumptionInput.value);

    if (isNaN(minVal) && isNaN(maxVal)) {
        showAlert('اخطار', 'لطفاً حداقل یا حداکثر مقدار مصرف را وارد کنید.', 'warning');
        log('warn', 'تلاش برای فیلتر مصرف بدون ورودی معتبر.');
        return;
    }

    // ابتدا فیلتر نام مشترک را اعمال می‌کنیم تا روی داده‌های فیلتر شده کار کنیم
    let dataToFilter = parsedData;
    const searchTerm = searchInput.value.toLowerCase();
    if (searchTerm !== '') {
        dataToFilter = parsedData.filter(row =>
            row['Customer name'] && row['Customer name'].toLowerCase().includes(searchTerm)
        );
    }


    filteredData = dataToFilter.filter(row => {
        const totalConsumption = row['Total Consumption [KWh]'];
        const isMinValid = isNaN(minVal) || totalConsumption >= minVal;
        const isMaxValid = isNaN(maxVal) || totalConsumption <= maxVal;
        return isMinValid && isMaxValid;
    });

    renderTable(filteredData);
    drawCharts(filteredData);
    log('info', `فیلتر بر اساس مصرف اعمال شد. حداقل: ${minVal || 'N/A'}, حداکثر: ${maxVal || 'N/A'}. تعداد نتایج: ${filteredData.length}`);
}


/**
 * پاک کردن فیلتر مصرف
 */
function clearConsumptionFilter() {
    minConsumptionInput.value = '';
    maxConsumptionInput.value = '';
    // پس از پاک کردن فیلتر مصرف، باید فیلتر نام مشترک را نیز دوباره اعمال کنیم
    filterDataByCustomerName();
    log('info', 'فیلتر مصرف پاک شد.');
}

/**
 * محاسبه میانگین مصرف برای بازه زمانی انتخابی
 */
function calculateAverageConsumptionForTimePeriod() {
    const timePeriod = timePeriodSelect.value;
    if (!timePeriod) {
        showAlert('اخطار', 'لطفاً یک بازه زمانی را انتخاب کنید.', 'warning');
        return;
    }

    let sumConsumption = 0;
    let count = 0;

    // پیدا کردن ستون‌های مربوط به بازه زمانی انتخاب شده
    const [startHour, endHour] = timePeriod.split('-').map(Number);
    const startColIndex = startHour * 4; // هر ساعت 4 تا 15 دقیقه دارد
    const endColIndex = endHour * 4;

    filteredData.forEach(row => {
        for (let i = startColIndex; i < endColIndex; i++) {
            const hour = String(Math.floor(i / 4)).padStart(2, '0');
            const minute = String((i % 4) * 15).padStart(2, '0');
            const nextMinute = String(((i % 4) + 1) * 15).padStart(2, '0');
            const colName = `${hour}:${minute} to ${hour}:${nextMinute} [KW]`;
            const value = row[colName];
            if (typeof value === 'number' && !isNaN(value)) {
                sumConsumption += value;
                count++;
            }
        }
    });

    const average = count > 0 ? sumConsumption / count : 0;
    timePeriodResultDiv.innerHTML = `<p>میانگین مصرف در بازه ${timePeriod}: <strong>${average.toLocaleString('fa-IR', { maximumFractionDigits: 2 })} کیلووات</strong></p>`;
    log('info', `میانگین مصرف در بازه ${timePeriod} محاسبه شد: ${average.toFixed(2)} KW.`);
}

/**
 * دانلود فایل لاگ
 */
function downloadLogFile() {
    const logContent = appLogs.map(log => `[${log.level.toUpperCase()}] ${log.timestamp}: ${log.message}`).join('\n');
    const blob = new Blob([logContent], { type: 'text/plain;charset=utf-8' });
    const date = new Date().toLocaleDateString('fa-IR').replace(/\//g, '-');
    saveAs(blob, `application_log_${date}.txt`);
    log('info', 'فایل لاگ دانلود شد.');
}

/**
 * اکسپورت جدول به Excel
 */
function exportTableToExcel() {
    if (filteredData.length === 0) {
        showAlert('اخطار', 'داده‌ای برای خروجی اکسل وجود ندارد.', 'warning');
        log('warn', 'تلاش برای اکسپورت اکسل بدون داده.');
        return;
    }

    const dataToExport = filteredData.map(row => {
        const newRow = {};
        DISPLAY_COLUMNS.forEach(col => {
            newRow[col] = row[col];
        });
        return newRow;
    });

    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "گزارش مصرف");
    const date = new Date().toLocaleDateString('fa-IR').replace(/\//g, '-');
    XLSX.writeFile(wb, `گزارش_مصرف_مشترکین_${date}.xlsx`);
    log('info', 'جدول به فایل اکسل اکسپورت شد.');
}

/**
 * اکسپورت نمودارها به صورت تصاویر
 */
async function exportChartsAsImages() {
    if (currentCharts.length === 0) {
        showAlert('اخطار', 'هیچ نموداری برای خروجی تصویر وجود ندارد.', 'warning');
        log('warn', 'تلاش برای اکسپورت تصاویر نمودار بدون وجود نمودار.');
        return;
    }

    Swal.fire({
        title: 'در حال آماده‌سازی تصاویر...',
        text: 'لطفاً منتظر بمانید.',
        allowOutsideClick: false,
        didOpen: () => {
            Swal.showLoading();
        }
    });

    const zip = new JSZip();
    const imagesFolder = zip.folder("charts");

    for (const chart of currentCharts) {
        try {
            const chartDataURL = chart.toBase64Image();
            // استخراج نام مشترک و تاریخ از عنوان نمودار
            const titleText = chart.options.plugins.title.text;
            const match = titleText.match(/مشترک: (.+) - تاریخ: (.+)/);
            let fileName = 'chart';
            if (match && match.length > 2) {
                const customerName = match[1].replace(/[<>:"/\\|?*]/g, '_'); // حذف کاراکترهای نامعتبر
                const date = match[2].replace(/[<>:"/\\|?*]/g, '_');
                fileName = `${customerName}_${date}`;
            } else {
                fileName = `chart_${new Date().getTime()}`; // نام فایل منحصر به فرد اگر اطلاعات یافت نشد
            }

            // تبدیل Data URL به Blob و افزودن به زیپ
            const blob = await fetch(chartDataURL).then(res => res.blob());
            imagesFolder.file(`${fileName}.png`, blob, { base64: true });
            log('info', `تصویر نمودار ${fileName} به فایل زیپ اضافه شد.`);

        } catch (error) {
            log('error', `خطا در ایجاد تصویر برای نمودار: ${error.message}`);
            Swal.fire('خطا', `خطا در ایجاد تصویر برای یکی از نمودارها: ${error.message}`, 'error');
            return;
        }
    }

    zip.generateAsync({ type: "blob" }).then(function(content) {
        const date = new Date().toLocaleDateString('fa-IR').replace(/\//g, '-');
        saveAs(content, `charts_images_${date}.zip`);
        Swal.close();
        showAlert('موفق', 'تصاویر نمودارها با موفقیت اکسپورت و دانلود شدند.', 'success');
        log('info', 'فایل زیپ تصاویر نمودارها دانلود شد.');
    }).catch(e => {
        log('error', `خطا در ایجاد فایل زیپ نمودارها: ${e.message}`);
        Swal.fire('خطا', `خطا در ایجاد فایل زیپ تصاویر: ${e.message}`, 'error');
    });
}

/**
 * اکسپورت به PDF
 */
async function exportToPdf() {
    if (filteredData.length === 0) {
        showAlert('اخطار', 'داده‌ای برای خروجی PDF وجود ندارد.', 'warning');
        log('warn', 'تلاش برای اکسپورت PDF بدون داده.');
        return;
    }

    Swal.fire({
        title: 'در حال ساخت PDF...',
        text: 'لطفاً منتظر بمانید. این فرآیند ممکن است کمی طول بکشد.',
        allowOutsideClick: false,
        didOpen: () => {
            Swal.showLoading();
        }
    });

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({
        orientation: 'landscape', // افقی
        unit: 'pt',
        format: 'a4'
    });

    const margin = 40;
    let yOffset = margin;
    const pageHeight = doc.internal.pageSize.height;
    const pageWidth = doc.internal.pageSize.width;

    // Load Vazirmatn font
    doc.addFont('Vazirmatn-Regular.ttf', 'Vazirmatn', 'normal');
    doc.setFont('Vazirmatn');
    doc.setFontSize(10);

    // Add Header
    doc.text('گزارش تحلیل مصرف برق مشترکین', pageWidth / 2, yOffset, { align: 'center' });
    yOffset += 20;

    // Add Table to PDF
    const tableHeaders = DISPLAY_COLUMNS.map(col => {
        // ترجمه سربرگ‌ها برای نمایش بهتر در PDF
        switch (col) {
            case '#': return 'ردیف';
            case 'Serial no.': return 'شماره سریال';
            case 'Date': return 'تاریخ';
            case 'Customer name': return 'نام مشترک';
            case 'Billing id': return 'شناسه قبض';
            case 'Customer id': return 'شناسه مشترک';
            case 'Contracted demand': return 'توان قراردادی';
            case 'Address': return 'آدرس';
            default: return col;
        }
    });

    const tableRows = filteredData.map((row, index) => {
        const newRow = [];
        DISPLAY_COLUMNS.forEach(col => {
            let cellValue = row[col];
            if (typeof cellValue === 'number') {
                cellValue = cellValue.toLocaleString('fa-IR', { maximumFractionDigits: 2 });
            }
            newRow.push(cellValue);
        });
        return newRow;
    });

    doc.autoTable({
        head: [tableHeaders],
        body: tableRows,
        startY: yOffset + 10,
        theme: 'grid',
        styles: {
            font: 'Vazirmatn',
            fontStyle: 'normal',
            halign: 'center',
            fontSize: 8,
            cellPadding: 2
        },
        headStyles: {
            fillColor: [30, 144, 255], // DodgerBlue
            textColor: 255,
            lineWidth: 0.5,
            lineColor: [255, 255, 255]
        },
        bodyStyles: {
            textColor: 50,
            lineWidth: 0.2,
            lineColor: [200, 200, 200]
        },
        alternateRowStyles: {
            fillColor: [240, 240, 240]
        },
        columnStyles: {
            // Apply specific styles for columns if needed
        },
        margin: { top: yOffset, right: margin, bottom: margin, left: margin },
        didDrawPage: function(data) {
            // Footer
            let str = "صفحه " + doc.internal.getNumberOfPages();
            doc.setFontSize(8);
            doc.text(str, pageWidth / 2, pageHeight - 10, { align: 'center' });
        }
    });

    yOffset = doc.autoTable.previous.finalY + 20;

    // Add Charts to PDF
    for (const chart of currentCharts) {
        if (yOffset + 200 > pageHeight - margin) { // Check if new page is needed for chart
            doc.addPage();
            yOffset = margin;
            // Add Header to new page
            doc.setFontSize(10);
            doc.text('گزارش تحلیل مصرف برق مشترکین (ادامه)', pageWidth / 2, yOffset, { align: 'center' });
            yOffset += 20;
        }

        try {
            const chartDataURL = chart.toBase64Image({
                format: 'image/jpeg',
                quality: 0.8
            });

            // Calculate image dimensions to fit within PDF
            const imgWidth = 500; // Example width, adjust as needed
            const imgHeight = (chart.canvas.height / chart.canvas.width) * imgWidth;

            doc.addImage(chartDataURL, 'JPEG', (pageWidth - imgWidth) / 2, yOffset, imgWidth, imgHeight);
            yOffset += imgHeight + 10; // Space after chart
            log('info', `نمودار به PDF اضافه شد.`);
        } catch (error) {
            log('error', `خطا در افزودن نمودار به PDF: ${error.message}`);
        }
    }

    const date = new Date().toLocaleDateString('fa-IR').replace(/\//g, '-');
    doc.save(`گزارش_مصرف_مشترکین_${date}.pdf`);
    Swal.close();
    showAlert('موفق', 'گزارش PDF با موفقیت ایجاد و دانلود شد.', 'success');
    log('info', 'گزارش PDF دانلود شد.');
}


// ====================================================================================================
// شنونده‌های رویداد (Event Listeners)
// ====================================================================================================

// انتخاب فایل اکسل
excelFile.addEventListener('change', (event) => {
    const file = event.target.files[0];
    if (file) {
        fileNameDisplay.textContent = file.name;
        log('info', `فایل انتخاب شد: ${file.name}`);
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array' });
            // پر کردن دراپ‌دان شیت‌ها
            sheetSelect.innerHTML = '';
            workbook.SheetNames.forEach(sheetName => {
                const option = document.createElement('option');
                option.value = sheetName;
                option.textContent = sheetName;
                sheetSelect.appendChild(option);
            });
            sheetSelect.disabled = false;
            // به صورت پیش‌فرض اولین شیت انتخاب شود
            if (workbook.SheetNames.length > 0) {
                sheetSelect.value = workbook.SheetNames[0];
            }
            elements.processBtn.disabled = false; // فعال کردن دکمه پردازش
            log('info', 'فایل اکسل بارگذاری شد. شیت‌ها پر شدند.');
        };
        reader.onerror = (e) => {
            log('error', `خطا در خواندن فایل: ${e.message}`);
            showAlert('خطا', 'خطا در خواندن فایل اکسل. لطفاً مجدداً تلاش کنید.', 'error');
        };
        reader.readAsArrayBuffer(file);
    } else {
        fileNameDisplay.textContent = 'فایل انتخاب نشده است.';
        sheetSelect.innerHTML = '<option value="">- Sheet1 -</option>';
        sheetSelect.disabled = true;
        elements.processBtn.disabled = true;
        log('info', 'انتخاب فایل لغو شد.');
    }
});

// پردازش فایل انتخاب شده
// پیدا کردن دکمه پردازش
// const processBtn = document.getElementById('processDataBtn');
if (elements.processBtn) {
    elements.processBtn.addEventListener('click', function() {
        // کدهای پردازش فایل اکسل
            if (!workbook) {
        showAlert('اخطار', 'لطفاً ابتدا یک فایل اکسل را انتخاب کنید.', 'warning');
        log('warn', 'تلاش برای پردازش بدون انتخاب فایل.');
        return;
    }
    const selectedSheetName = sheetSelect.value;
    if (!selectedSheetName) {
        showAlert('اخطار', 'لطفاً یک شیت را انتخاب کنید.', 'warning');
        log('warn', 'تلاش برای پردازش بدون انتخاب شیت.');
        return;
    }

    const ws = workbook.Sheets[selectedSheetName];
    const jsonData = XLSX.utils.sheet_to_json(ws);

    if (jsonData.length === 0) {
        showAlert('اخطار', 'شیت انتخاب شده حاوی داده‌ای نیست.', 'warning');
        log('warn', `شیت "${selectedSheetName}" خالی است.`);
        return;
    }

    const headers = XLSX.utils.sheet_to_json(ws, { header: 1 })[0];
    log('error',  headers);
    if (!validateHeaders(headers)) {
        log('error', 'هدرهای فایل اکسل نامعتبر هستند. پردازش متوقف شد.');
        return;
    }

    parsedData = processExcelData(jsonData);
    filteredData = [...parsedData]; // در ابتدا، داده‌های فیلتر شده همان داده‌های پردازش شده هستند

    renderTable(filteredData);
    drawCharts(filteredData);

    // فعال کردن دکمه‌های اکسپورت و فیلتر
    exportPdfBtn.disabled = false;
    exportExcelBtn.disabled = false;
    exportChartsAsImagesBtn.disabled = false;
    searchInput.disabled = false;
    filterCustomerBtn.disabled = false;
    clearFilterBtn.disabled = false;
    minConsumptionInput.disabled = false;
    maxConsumptionInput.disabled = false;
    // filterConsumptionBtn.disabled = false;
    clearConsumptionFilterBtn.disabled = false;
    timePeriodSelect.disabled = false;
    calculateTimePeriodBtn.disabled = false;
    renderAllChartsBtn.disabled = false;

    showAlert('موفق', 'فایل اکسل با موفقیت پردازش شد و نمودارها رسم شدند.', 'success');
    log('info', 'پردازش موفقیت‌آمیز فایل اکسل و رسم نمودارها.');
    });
} else {
    console.error('دکمه پردازش یافت نشد!');
    // نمایش پیام به کاربر
    if (typeof Swal !== 'undefined') {
        Swal.fire('خطا', 'دکمه پردازش در صفحه یافت نشد!', 'error');
    }
}
// شنونده برای دکمه فیلتر نام مشترک
if (filterCustomerBtn) {
    filterCustomerBtn.addEventListener('click', filterDataByCustomerName);
}

// شنونده برای دکمه پاک کردن فیلتر نام مشترک
if (clearFilterBtn) {
    clearFilterBtn.addEventListener('click', clearCustomerFilter);
}

// شنونده برای دکمه فیلتر مصرف
// if (filterConsumptionBtn) {
    // filterConsumptionBtn.addEventListener('click', filterDataByConsumption);
// }

// شنونده برای دکمه پاک کردن فیلتر مصرف
if (clearConsumptionFilterBtn) {
    clearConsumptionFilterBtn.addEventListener('click', clearConsumptionFilter);
}

// شنونده برای دکمه محاسبه میانگین مصرف در بازه زمانی
if (calculateTimePeriodBtn) {
    calculateTimePeriodBtn.addEventListener('click', calculateAverageConsumptionForTimePeriod);
}

// شنونده‌های دکمه‌های اکسپورت
if (exportPdfBtn) {
    // بررسی وجود jsPDF و html2canvas
    if (typeof jspdf === 'undefined' || typeof html2canvas === 'undefined') {
        exportPdfBtn.disabled = true;
        log('error', 'عدم بارگذاری کتابخانه‌های مورد نیاز (jspdf, html2canvas) برای خروجی PDF.');
        Swal.fire('خطا', 'عدم بارگذاری کتابخانه‌های مورد نیاز (jspdf, html2canvas) برای خروجی PDF.', 'error');
    } else {
        exportPdfBtn.addEventListener('click', exportToPdf);
        // Add font for jsPDF
        // This is a placeholder for adding a font file that supports Persian.
        // In a real scenario, you'd need to load the font data.
        // For example:
        // doc.addFont('path/to/Vazirmatn-Regular.ttf', 'Vazirmatn', 'normal');
        // doc.setFont('Vazirmatn');
    }
}

if (exportExcelBtn) {
    // بررسی وجود XLSX
    if (typeof XLSX === 'undefined') {
        exportExcelBtn.disabled = true;
        log('error', 'عدم بارگذاری کتابخانه مورد نیاز (XLSX.js) برای خروجی اکسل.');
        Swal.fire('خطا', 'عدم بارگذاری کتابخانه مورد نیاز (XLSX.js) برای خروجی اکسل.', 'error');
    } else {
        exportExcelBtn.addEventListener('click', exportTableToExcel);
    }
}

// برای اکسپورت تصاویر نمودارها به JSZip و FileSaver.js نیاز داریم
if (exportChartsAsImagesBtn) {
    if (typeof JSZip === 'undefined' || typeof saveAs === 'undefined') {
        exportChartsAsImagesBtn.disabled = true;
        // تلاش برای بارگذاری دینامیک کتابخانه‌ها
        log('warn', 'کتابخانه‌های JSZip یا FileSaver.js بارگذاری نشده‌اند. تلاش برای بارگذاری.');
        const script = document.createElement('script');
        script.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
        script.onload = () => {
            log('info', 'JSZip با موفقیت بارگذاری شد.');
            const scriptSaveAs = document.createElement('script');
            scriptSaveAs.src = "https://cdnjs.cloudflare.com/ajax/libs/FileSaver.js/2.0.5/FileSaver.min.js";
            scriptSaveAs.onload = () => {
                log('info', 'FileSaver.js با موفقیت بارگذاری شد.');
                exportChartsAsImagesBtn.disabled = false;
                exportChartsAsImagesBtn.addEventListener('click', exportChartsAsImages);
            };
            scriptSaveAs.onerror = () => {
                log('error', 'خطا در بارگذاری کتابخانه FileSaver.js. اکسپورت تصویر نمودار کار نخواهد کرد.');
                Swal.fire('خطا', 'عدم بارگذاری کتابخانه مورد نیاز (FileSaver.js) برای خروجی تصویر نمودارها.', 'error');
            };
            document.head.appendChild(scriptSaveAs);
        };
        script.onerror = () => {
            log('error', 'خطا در بارگذاری کتابخانه JSZip. اکسپورت تصویر نمودار کار نخواهد کرد.');
            Swal.fire('خطا', 'عدم بارگذاری کتابخانه مورد نیاز (JSZip) برای خروجی تصویر نمودارها.', 'error');
        };
        document.head.appendChild(script);
    } else {
        exportChartsAsImagesBtn.addEventListener('click', exportChartsAsImages);
    }
}
if (exportLogFileBtn) {
    exportLogFileBtn.addEventListener('click', downloadLogFile);
}
//------------------------------------------------------
// Render All Charts Button Listener (if exists in HTML)
if (renderAllChartsBtn) {
    renderAllChartsBtn.addEventListener('click', () => {
        if (filteredData.length > 0) {
            drawCharts(filteredData); // Re-draw all charts based on current filtered data
            Swal.fire('نمودارها بازسازی شد', 'تمام نمودارها مجدداً رسم شدند.', 'info');
            log('info', 'تمام نمودارها به صورت دستی بازسازی شدند.');
        } else {
            Swal.fire('اخطار', 'داده‌ای برای رسم نمودار وجود ندارد. لطفاً ابتدا فایل را پردازش کنید.', 'warning');
            log('warn', 'تلاش برای بازسازی نمودارها بدون داده.');
        }
    });
}

/**
 * تابع اصلی پردازش داده‌ها از فایل اکسل
 * این تابع پس از انتخاب شیت و کلیک دکمه "پردازش" فراخوانی می‌شود.
 */
async function processData() {
    const sheetSelect = document.getElementById('sheetSelect');
    const resultsTableBody = document.querySelector('#resultsTable tbody');
    const chartsContainer = document.getElementById('chartsContainer');
    const noChartsMessage = document.getElementById('noChartsMessage')  || { style: { display: '' } };
    const morningStartHour = document.getElementById('morningStartHour');
    const morningStartMinute = document.getElementById('morningStartMinute');
    const morningEndHour = document.getElementById('morningEndHour');
    const morningEndMinute = document.getElementById('morningEndMinute');
    const eveningStartHour = document.getElementById('eveningStartHour');
    const eveningStartMinute = document.getElementById('eveningStartMinute');
    const eveningEndHour = document.getElementById('eveningEndHour');
    const eveningEndMinute = document.getElementById('eveningEndMinute');
    const morningCalcType = document.getElementById('morningCalcType');
    const eveningCalcType = document.getElementById('eveningCalcType');
    const chkEvening = document.getElementById('chkEvening');
    const txtEvening = document.getElementById('txtEvening');
    const chkReduction = document.getElementById('chkReduction');
    const txtReduction = document.getElementById('txtReduction');
    const exportExcelBtn = document.getElementById('exportExcelBtn');
    const exportPdfBtn = document.getElementById('exportPdfBtn');
    const exportChartsAsImagesBtn = document.getElementById('exportChartsAsImagesBtn');


    if (!workbook) {
        Swal.fire('خطا', 'لطفاً ابتدا یک فایل اکسل انتخاب کنید.', 'error');
        return;
    }

    const selectedSheetName = sheetSelect.value || workbook.SheetNames[0];
    if (!selectedSheetName) {
        Swal.fire('خطا', 'شیتی برای پردازش یافت نشد.', 'error');
        return;
    }

    showProgress(25, 'در حال پردازش داده‌ها...');

    if (resultsTableBody) resultsTableBody.innerHTML = '';
    destroyCharts();
    if (noChartsMessage) noChartsMessage.style.display = 'none';

    const worksheet = workbook.Sheets[selectedSheetName];
    const jsonSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });

    const headers = jsonSheet[0];
    console.log("سربرگ‌ها از اکسل:", headers);
    const dataRows = jsonSheet.slice(1);
    console.log("ردیف‌های داده (dataRows):", dataRows);

    parsedData = []; // پاکسازی آرایه داده‌های پردازش‌شده برای شروع مجدد

    const getMinutes = (h, m) => parseInt(h) * 60 + parseInt(m);
    // ترتیب ساعت و دقیقه در ورودی های UI تصحیح شد
    const morningStart = getMinutes(morningStartHour.value, morningStartMinute.value);
    const morningEnd = getMinutes(morningEndHour.value, morningEndMinute.value);
    const eveningStart = getMinutes(eveningStartHour.value, eveningStartMinute.value);
    const eveningEnd = getMinutes(eveningEndHour.value, eveningEndMinute.value);


    const bodyNumberColIndex = headers.indexOf('Serial no.');
    const customerNameColIndex = headers.indexOf('Customer name');
    const billIdColIndex = headers.indexOf('Billing id');
    const addressColIndex = headers.indexOf('Address');
    const subscriptionNumberColIndex = headers.indexOf('Customer id');
    const demandColumnIndex = headers.indexOf('Contracted demand');

    if (bodyNumberColIndex === -1 || customerNameColIndex === -1 || billIdColIndex === -1 ||
        addressColIndex === -1 || subscriptionNumberColIndex === -1 || demandColumnIndex === -1) {
        Swal.fire({
            icon: 'error',
            title: 'خطای ساختار فایل',
            html: 'یکی از ستون‌های ضروری (مانند "Serial no.", "Customer name", "Billing id", "Address", "Customer id", "Contracted demand") در فایل اکسل یافت نشد. <br> لطفاً مطمئن شوید که سربرگ‌ها صحیح هستند.'
        });
        showProgress(0);
        return;
    }

    const minEveningLoad = chkEvening && chkEvening.checked ? parseFloat(txtEvening.value) : -Infinity;
    // Changed to maxReductionPercent and default to Infinity
    const maxReductionPercent = chkReduction && chkReduction.checked ? parseFloat(txtReduction.value) : Infinity;

    showProgress(50, 'در حال تحلیل بارهای مشترکین...');

    for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
        if (!row || row.length === 0 || row[bodyNumberColIndex] === undefined || row[bodyNumberColIndex] === null) {
            console.warn(`ردیف ${i + 2} به دلیل خالی بودن یا نداشتن شماره بدنه، نادیده گرفته شد.`);
            continue;
        }

        const customerInfo = {
            // ID با bodyNumber ساخته شد تا پیدا کردن نمودار آسان‌تر باشد
            id: `customer-${row[bodyNumberColIndex]}`,
            rowNum: i + 2,
            bodyNumber: row[bodyNumberColIndex],
            customerName: row[customerNameColIndex],
            billId: row[billIdColIndex],
            address: row[addressColIndex],
            subscriptionNumber: row[subscriptionNumberColIndex],
            contractDemand: parseFloat(row[demandColumnIndex]) || 0
        };

        const loadProfile = [];
        const timeLabels = []; // This will now store H1, H2, ... H24

        // ایندکس آخرین ستون اطلاعات ثابت + 1 (فرض بر این است که ستون‌های بار بلافاصله بعد از اینها شروع می‌شوند)
        const firstLoadColumnIndex = Math.max(
            headers.indexOf('#'), // اضافه کردن # به لیست تا مطمئن شویم از اولین ستون‌های بار شروع می‌کنیم
            bodyNumberColIndex, customerNameColIndex, billIdColIndex,
            addressColIndex, subscriptionNumberColIndex, demandColumnIndex
        ) + 1;


        for (let j = firstLoadColumnIndex; j < headers.length; j++) {
            const header = headers[j];
            // اصلاح شده برای تطبیق با فرمت "HH:MM to HH:MM [KW]"
            const timeMatch = String(header).match(/^(\d{2}:\d{2}) to \d{2}:\d{2} \[KW\]$/);

            if (timeMatch && timeMatch[1]) {
                const timeString = timeMatch[1];
                const [h, m] = timeString.split(':').map(Number);
                const timeInMinutes = h * 60 + m;
                const loadValue = parseFloat(row[j]);
                if (!isNaN(loadValue)) {
                    loadProfile.push({ timeInMinutes, load: loadValue });
                    // Convert HH:MM to H1, H2, ... H24
                    timeLabels.push(`H${h + 1}`); // Assuming h is 0-23
                } else {
                    console.warn(`مقدار بار نامعتبر در ردیف ${i + 2}, ستون ${header}: ${row[j]}`);
                }
            } else {
                // این خط برای دیباگینگ مفید است، نشان می‌دهد کدام سربرگ‌ها نادیده گرفته می‌شوند
                // console.warn(`سربرگ "${header}" به عنوان ستون زمان بار معتبر شناسایی نشد.`);
            }
        }

        if (loadProfile.length === 0) {
            console.warn(`ردیف ${i + 2} (${customerInfo.bodyNumber}) به دلیل پروفایل بار خالی (عدم یافتن ستون‌های زمانی معتبر)، نادیده گرفته شد.`);
            continue;
        }

        const morningLoads = loadProfile.filter(item =>
            item.timeInMinutes >= morningStart && item.timeInMinutes <= morningEnd
        ).map(item => item.load);

        let morningLoad = 0;
        if (morningLoads.length > 0) {
            if (morningCalcType.value === 'avg') morningLoad = morningLoads.reduce((a, b) => a + b, 0) / morningLoads.length;
            else if (morningCalcType.value === 'max') morningLoad = Math.max(...morningLoads);
            else if (morningCalcType.value === 'min') morningLoad = Math.min(...morningLoads);
        }

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
            loadProfileData: loadProfile.map(item => item.load),
            timeLabels: timeLabels
        };

        const passesEveningFilter = !chkEvening || !chkEvening.checked || (parseFloat(customerResult.eveningLoad) >= minEveningLoad);
        // Changed comparison for max reduction percentage
        const passesReductionFilter = !chkReduction || !chkReduction.checked || (parseFloat(customerResult.reductionPercent) <= maxReductionPercent);

        if (passesEveningFilter && passesReductionFilter) {
            parsedData.push(customerResult);
            console.log("مشتری با موفقیت اضافه شد:", customerResult);
        } else {
            console.log(`مشتری ${customerInfo.bodyNumber} به دلیل عدم تطابق با فیلترها اضافه نشد.`, {
                passesEveningFilter,
                passesReductionFilter,
                eveningLoad: customerResult.eveningLoad,
                minEveningLoad,
                reductionPercent: customerResult.reductionPercent,
                maxReductionPercent // Changed variable name
            });
        }
    }

    console.log("آرایه نهایی parsedData پس از پردازش:", parsedData);

    showProgress(75, 'در حال نمایش نتایج...');
    displayResults();
    drawCharts();
    showProgress(100, 'پردازش کامل شد.');

    // اطمینان از فعال شدن دکمه‌های خروجی تنها در صورت وجود داده
    if (parsedData.length > 0) {
        if (exportExcelBtn) exportExcelBtn.disabled = false;
        if (exportPdfBtn) exportPdfBtn.disabled = false;
        if (exportChartsAsImagesBtn) exportChartsAsImagesBtn.disabled = false;
    } else {
        if (exportExcelBtn) exportExcelBtn.disabled = true;
        if (exportPdfBtn) exportPdfBtn.disabled = true;
        if (exportChartsAsImagesBtn) exportChartsAsImagesBtn.disabled = true;
    }
}

/**
 * تابع برای نمایش نتایج پردازش شده در جدول HTML
 */
function displayResults() {
    const resultsTableBody = document.querySelector('#resultsTable tbody');

    if (resultsTableBody) {
        resultsTableBody.innerHTML = '';

        if (parsedData.length === 0) {
            const row = resultsTableBody.insertRow();
            const cell = row.insertCell();
            cell.colSpan = 13; // تعداد ستون‌ها
            cell.textContent = 'هیچ داده‌ای بر اساس فیلترهای اعمال شده یافت نشد.';
            cell.style.textAlign = 'center';
            console.log("هیچ داده‌ای برای نمایش در جدول وجود ندارد.");
            return;
        }

        parsedData.forEach((customer) => {
            const row = resultsTableBody.insertRow();
            row.dataset.customerId = customer.id; // برای دسترسی آسان به ID مشترک هنگام حذف

            row.insertCell().textContent = customer.rowNum;
            row.insertCell().textContent = customer.bodyNumber;
            row.insertCell().textContent = customer.customerName;
            row.insertCell().textContent = customer.billId;
            row.insertCell().textContent = customer.address;
            row.insertCell().textContent = customer.subscriptionNumber;
            row.insertCell().textContent = customer.contractDemand;
            row.insertCell().textContent = customer.morningLoad;
            row.insertCell().textContent = customer.eveningLoad;
            row.insertCell().textContent = customer.reductionKW;
            row.insertCell().textContent = customer.reductionPercent;

            // اضافه کردن دکمه "مشاهده نمودار"
            const chartCell = row.insertCell();
            const viewChartBtn = document.createElement('button');
            viewChartBtn.textContent = 'نمودار';
            viewChartBtn.className = 'btn btn-info btn-sm';
            viewChartBtn.onclick = () => scrollToChart(customer.bodyNumber);
            chartCell.appendChild(viewChartBtn);

            // اضافه کردن دکمه "حذف"
            const deleteCell = row.insertCell();
            const deleteBtn = document.createElement('button');
            deleteBtn.textContent = 'حذف';
            deleteBtn.className = 'btn btn-danger btn-sm';
            // حذف مستقیم بدون سوال
            deleteBtn.onclick = () => deleteCustomerRow(customer.id);
            deleteCell.appendChild(deleteBtn);
        });
        console.log(`تعداد ${parsedData.length} مشتری در جدول نمایش داده شد.`);
    }
}

