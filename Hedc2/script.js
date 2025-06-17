// دسترسی به عناصر DOM
const fileInput = document.getElementById('excelFileInput');
const sheetSelect = document.getElementById('sheetSelect');
const processBtn = document.getElementById('processBtn');
const resultsTableBody = document.getElementById('resultsTableBody');
const progressContainer = document.getElementById('progressContainer');
const progressBar = document.getElementById('progressBar');
const progressLabel = document.getElementById('progressLabel');
const chartsContainer = document.getElementById('chartsContainer');
const noChartsMessage = document.getElementById('noChartsMessage');
const exportExcelBtn = document.getElementById('exportExcelBtn');
const exportPdfBtn = document.getElementById('exportPdfBtn');

// فیلترها و تنظیمات محاسبه
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

let workbook; // متغیر سراسری برای نگهداری شیء ورک‌بوک اکسل
let parsedData = []; // متغیر سراسری برای نگهداری داده‌های پردازش‌شده مشترکین
let currentCharts = []; // آرایه‌ای برای نگهداری نمونه‌های Chart.js (برای پاکسازی)

// تابع نمایش پیشرفت (Progress Bar)
function showProgress(percent, label = 'در حال بارگذاری...') {
    progressBar.style.width = percent + '%';
    progressBar.setAttribute('aria-valuenow', percent);
    progressLabel.textContent = label;
    if (percent === 0 || percent === 100) {
        progressContainer.style.display = 'none';
    } else {
        progressContainer.style.display = 'block';
    }
}

// تابع برای از بین بردن نمودارهای قبلی Chart.js
function destroyCharts() {
    currentCharts.forEach(chart => {
        if (chart) {
            chart.destroy();
        }
    });
    currentCharts = []; // پاکسازی آرایه
    chartsContainer.innerHTML = ''; // پاک کردن محتوای کانتینر نمودارها
}

// گوش دادن به رویداد تغییر فایل ورودی
fileInput.addEventListener('change', async (event) => {
    const file = event.target.files[0];
    if (!file) {
        return;
    }

    showProgress(10, 'در حال خواندن فایل...');

    const reader = new FileReader();
    reader.onload = async (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array' });
            console.log("ورک‌بوک با موفقیت بارگذاری شد:", workbook); // لاگ: ورک‌بوک بارگذاری شده

            // پر کردن دراپ‌داون شیت‌ها
            sheetSelect.innerHTML = ''; // پاکسازی گزینه‌های قبلی
            workbook.SheetNames.forEach(sheetName => {
                const option = document.createElement('option');
                option.value = sheetName;
                option.textContent = sheetName;
                sheetSelect.appendChild(option);
            });
            sheetSelect.disabled = false;
            processBtn.disabled = false;
            showProgress(100, 'فایل آماده پردازش است.');
        } catch (error) {
            console.error("خطا در بارگذاری فایل اکسل:", error); // لاگ: خطای بارگذاری
            Swal.fire('خطا', 'فایل اکسل نامعتبر است یا در خواندن آن مشکلی پیش آمده.', 'error');
            showProgress(0);
        }
    };
    reader.readAsArrayBuffer(file);
});

// گوش دادن به رویداد کلیک دکمه پردازش
processBtn.addEventListener('click', processData);
exportExcelBtn.addEventListener('click', exportToExcel);
exportPdfBtn.addEventListener('click', exportToPdf);

// تابع اصلی پردازش داده‌ها
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

    // پاکسازی نتایج و نمودارهای قبلی
    resultsTableBody.innerHTML = '';
    destroyCharts(); // از بین بردن نمونه‌های نمودار قبلی
    chartsContainer.innerHTML = ''; // پاک کردن محتوای کانتینر نمودار
    noChartsMessage.style.display = 'none'; // مخفی کردن پیام "نموداری وجود ندارد"

    // خواندن داده‌ها از شیت انتخاب‌شده
    const worksheet = workbook.Sheets[selectedSheetName];
    // تبدیل شیت به JSON، با رد کردن سربرگ (header:1) برای دریافت آرایه‌ای از آرایه‌ها
    // استفاده از raw: true برای دریافت مقادیر خام، نه فرمت‌شده
    const jsonSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });

    // فرض می‌کنیم ردیف اول سربرگ‌ها و داده‌ها از ردیف دوم شروع می‌شوند
    const headers = jsonSheet[0];
    console.log("سربرگ‌ها از اکسل:", headers); // لاگ: سربرگ‌های خوانده شده
    const dataRows = jsonSheet.slice(1);
    console.log("ردیف‌های داده (dataRows):", dataRows); // لاگ: ردیف‌های داده خام

    parsedData = []; // پاکسازی داده‌های قبلی قبل از شروع پردازش جدید

    // یک بار محاسبه مرزهای زمانی (صبح و شب)
    const getMinutes = (h, m) => parseInt(h) * 60 + parseInt(m);
    const morningStart = getMinutes(morningStartHour.value, morningStartMinute.value);
    const morningEnd = getMinutes(morningEndHour.value, morningEndMinute.value);
    const eveningStart = getMinutes(eveningStartHour.value, eveningStartMinute.value);
    const eveningEnd = getMinutes(eveningEndHour.value, eveningEndMinute.value);

    // شناسایی ستون‌های مربوطه بر اساس نام سربرگ
    const bodyNumberColIndex = headers.indexOf('Serial no.');
    const customerNameColIndex = headers.indexOf('Customer name');
    const billIdColIndex = headers.indexOf('Billing id');
    const addressColIndex = headers.indexOf('Address');
    const subscriptionNumberColIndex = headers.indexOf('Customer id');
    const demandColumnIndex = headers.indexOf('Contracted demand');

    // بررسی وجود ستون‌های ضروری
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

    // مقادیر فیلترها را از UI دریافت کنید
    const minEveningLoad = chkEvening.checked ? parseFloat(txtEvening.value) : -Infinity;
    const minReductionPercent = chkReduction.checked ? parseFloat(txtReduction.value) : -Infinity;

    progressLabel.textContent = 'در حال تحلیل بارهای مشترکین...';
    showProgress(50);

    // حلقه بر روی هر ردیف داده
    for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
        // از ردیف‌های خالی یا ردیف‌هایی که شبیه ردیف داده نیستند (مثلاً شماره بدنه ندارند) بگذرید
        if (!row || row.length === 0 || row[bodyNumberColIndex] === undefined || row[bodyNumberColIndex] === null) {
            console.warn(`ردیف ${i + 2} به دلیل خالی بودن یا نداشتن شماره بدنه، نادیده گرفته شد.`); // لاگ: ردیف نادیده گرفته شده
            continue;
        }

        const customerInfo = {
            rowNum: i + 2, // شماره ردیف اصلی در اکسل (با فرض اینکه سربرگ ردیف 1 است)
            bodyNumber: row[bodyNumberColIndex],
            customerName: row[customerNameColIndex],
            billId: row[billIdColIndex],
            address: row[addressColIndex],
            subscriptionNumber: row[subscriptionNumberColIndex],
            contractDemand: parseFloat(row[demandColumnIndex]) || 0 // اگر نامعتبر بود، 0 در نظر گرفته شود
        };

        const loadProfile = []; // برای ذخیره { timeInMinutes, loadValue } برای هر ردیف
        const timeLabels = []; // برای ذخیره برچسب‌های زمانی برای نمودار

        // شناسایی اولین ستون داده بار
        // این روش فرض می‌کند که ستون‌های بار بلافاصله بعد از آخرین ستون اطلاعات ثابت شروع می‌شوند.
        // اگر ساختار فایل شما ثابت است و ستون‌های بار همیشه از یک ایندکس مشخص شروع می‌شوند (مثلاً ایندکس 7)،
        // می‌توانید به جای این محاسبات، مستقیماً `const firstLoadColumnIndex = 7;` را قرار دهید.
        const firstLoadColumnIndex = Math.max(
            bodyNumberColIndex, customerNameColIndex, billIdColIndex,
            addressColIndex, subscriptionNumberColIndex, demandColumnIndex
        ) + 1;

        // حلقه بر روی ستون‌های داده بار
        for (let j = firstLoadColumnIndex; j < headers.length; j++) {
            const header = headers[j];
            // بررسی کنید که آیا سربرگ شبیه یک زمان است (مثلاً "00:00", "00:15")
            if (typeof header === 'string' && header.match(/^\d{2}:\d{2}$/)) {
                const [h, m] = header.split(':').map(Number);
                const timeInMinutes = h * 60 + m;
                const loadValue = parseFloat(row[j]);
                if (!isNaN(loadValue)) {
                    loadProfile.push({ timeInMinutes, load: loadValue });
                    timeLabels.push(header); // استفاده از رشته اصلی برای برچسب
                } else {
                    console.warn(`مقدار بار نامعتبر در ردیف ${i + 2}, ستون ${header}: ${row[j]}`); // لاگ: مقدار بار نامعتبر
                }
            } else {
                // اگر سربرگ زمان نیست، آن را نادیده بگیرید یا لاگ کنید
                // console.log(`سربرگ ${header} در ستون ${j} شبیه زمان نیست، نادیده گرفته شد.`);
            }
        }

        if (loadProfile.length === 0) {
            console.warn(`ردیف ${i + 2} (${customerInfo.bodyNumber}) به دلیل پروفایل بار خالی، نادیده گرفته شد. این به این معنی است که هیچ ستون زمانی معتبری بعد از ستون‌های اطلاعات ثابت یافت نشد.`); // لاگ: ردیف با پروفایل بار خالی
            continue; // اگر داده باری یافت نشد، از این مشتری بگذرید
        }

        // محاسبه بار صبح
        const morningLoads = loadProfile.filter(item =>
            item.timeInMinutes >= morningStart && item.timeInMinutes <= morningEnd
        ).map(item => item.load);

        let morningLoad = 0;
        if (morningLoads.length > 0) {
            if (morningCalcType.value === 'avg') morningLoad = morningLoads.reduce((a, b) => a + b, 0) / morningLoads.length;
            else if (morningCalcType.value === 'max') morningLoad = Math.max(...morningLoads);
            else if (morningCalcType.value === 'min') morningLoad = Math.min(...morningLoads);
        }

        // محاسبه بار شب
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
            loadProfileData: loadProfile.map(item => item.load), // فقط مقادیر بار برای داده‌های نمودار
            timeLabels: timeLabels // برچسب‌های زمانی برای محور X نمودار
        };

        // اعمال فیلترها
        const passesEveningFilter = !chkEvening.checked || (parseFloat(customerResult.eveningLoad) >= minEveningLoad);
        const passesReductionFilter = !chkReduction.checked || (parseFloat(customerResult.reductionPercent) >= minReductionPercent);

        if (passesEveningFilter && passesReductionFilter) {
            parsedData.push(customerResult);
            console.log("مشتری با موفقیت اضافه شد:", customerResult); // لاگ: مشتری اضافه شده به parsedData
        } else {
            console.log(`مشتری ${customerInfo.bodyNumber} به دلیل عدم تطابق با فیلترها اضافه نشد.`, {
                passesEveningFilter,
                passesReductionFilter,
                eveningLoad: customerResult.eveningLoad,
                minEveningLoad,
                reductionPercent: customerResult.reductionPercent,
                minReductionPercent
            }); // لاگ: مشتری رد شده توسط فیلتر
        }
    } // پایان حلقه for (let i = 0; i < dataRows.length; i++)

    console.log("آرایه نهایی parsedData پس از پردازش:", parsedData); // لاگ: وضعیت نهایی parsedData

    showProgress(75, 'در حال نمایش نتایج...');
    displayResults();
    drawCharts();
    showProgress(100, 'پردازش کامل شد.');

    exportExcelBtn.disabled = false;
    exportPdfBtn.disabled = false;
}

// تابع برای نمایش نتایج در جدول
function displayResults() {
    resultsTableBody.innerHTML = ''; // پاکسازی جدول قبل از افزودن نتایج جدید

    if (parsedData.length === 0) {
        const row = resultsTableBody.insertRow();
        const cell = row.insertCell();
        cell.colSpan = 8; // تعداد ستون‌ها
        cell.textContent = 'هیچ داده‌ای بر اساس فیلترهای اعمال شده یافت نشد.';
        cell.style.textAlign = 'center';
        console.log("هیچ داده‌ای برای نمایش در جدول وجود ندارد."); // لاگ: جدول خالی
        return;
    }

    parsedData.forEach(customer => {
        const row = resultsTableBody.insertRow();
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
        viewChartBtn.onclick = () => scrollToChart(customer.bodyNumber); // اسکرول به نمودار مربوطه
        chartCell.appendChild(viewChartBtn);
    });
    console.log(`تعداد ${parsedData.length} مشتری در جدول نمایش داده شد.`); // لاگ: تعداد ردیف‌های جدول
}

// تابع برای رسم نمودارها
function drawCharts() {
    destroyCharts(); // ابتدا هر نمودار موجود را پاک کنید

    if (parsedData.length === 0) {
        noChartsMessage.style.display = 'block'; // نمایش پیام "نموداری وجود ندارد"
        console.log("هیچ داده‌ای برای رسم نمودار وجود ندارد."); // لاگ: نمودارها خالی
        return;
    }

    noChartsMessage.style.display = 'none';

    parsedData.forEach(customer => {
        const chartId = `chart-${customer.bodyNumber}`;
        const chartDiv = document.createElement('div');
        chartDiv.className = 'chart-container';
        chartDiv.id = `chart-div-${customer.bodyNumber}`; // برای اسکرول کردن به آن
        chartDiv.innerHTML = `
            <h3>نمودار بار مصرفی مشترک: ${customer.customerName} (شماره بدنه: ${customer.bodyNumber})</h3>
            <canvas id="${chartId}"></canvas>
            <hr>
        `;
        chartsContainer.appendChild(chartDiv);

        const ctx = document.getElementById(chartId).getContext('2d');
        const newChart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: customer.timeLabels,
                datasets: [{
                    label: 'بار مصرفی (KW)',
                    data: customer.loadProfileData,
                    borderColor: 'rgb(75, 192, 192)',
                    tension: 0.1,
                    fill: false
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                    x: {
                        title: {
                            display: true,
                            text: 'زمان'
                        }
                    },
                    y: {
                        title: {
                            display: true,
                            text: 'بار (KW)'
                        },
                        beginAtZero: true
                    }
                },
                plugins: {
                    tooltip: {
                        mode: 'index',
                        intersect: false,
                    },
                    legend: {
                        display: true,
                        position: 'top',
                    }
                }
            }
        });
        currentCharts.push(newChart); // اضافه کردن نمونه نمودار به آرایه
        console.log(`نمودار برای مشتری ${customer.bodyNumber} رسم شد.`); // لاگ: نمودار رسم شده
    });
    console.log(`تعداد ${currentCharts.length} نمودار رسم شد.`); // لاگ: تعداد کل نمودارها
}

// تابع برای اسکرول کردن به نمودار خاص
function scrollToChart(bodyNumber) {
    const chartDiv = document.getElementById(`chart-div-${bodyNumber}`);
    if (chartDiv) {
        chartDiv.scrollIntoView({ behavior: 'smooth', block: 'start' });
        console.log(`به نمودار مشتری ${bodyNumber} اسکرول شد.`); // لاگ: اسکرول
    } else {
        console.warn(`نمودار مشتری ${bodyNumber} یافت نشد.`); // لاگ: نمودار یافت نشد
    }
}

// تابع برای خروجی اکسل
function exportToExcel() {
    if (parsedData.length === 0) {
        Swal.fire('هشدار', 'داده‌ای برای خروجی اکسل وجود ندارد.', 'warning');
        return;
    }

    const wsData = [
        ["شماره بدنه", "نام مشترک", "شناسه قبض", "آدرس مشترک", "شماره اشتراک", "دیماند قراردادی (KW)", "بار صبح (KW)", "بار شب (KW)", "کاهش بار (KW)", "درصد کاهش بار (%)"]
    ];
    parsedData.forEach(customer => {
        wsData.push([
            customer.bodyNumber,
            customer.customerName,
            customer.billId,
            customer.address,
            customer.subscriptionNumber,
            customer.contractDemand,
            customer.morningLoad,
            customer.eveningLoad,
            customer.reductionKW,
            customer.reductionPercent
        ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "نتایج");
    XLSX.writeFile(wb, "نتایج_تحلیل_بار.xlsx");
    console.log("خروجی اکسل ایجاد شد."); // لاگ: خروجی اکسل
}

// تابع برای خروجی PDF
function exportToPdf() {
    if (parsedData.length === 0) {
        Swal.fire('هشدار', 'داده‌ای برای خروجی PDF وجود ندارد.', 'warning');
        return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'pt', 'a4'); // 'p' for portrait, 'pt' for points, 'a4' for size

    // برای پشتیبانی از فارسی
    doc.addFont('Amiri-Regular.ttf', 'Amiri', 'normal');
    doc.setFont('Amiri');

    const tableColumn = ["شماره بدنه", "نام مشترک", "شناسه قبض", "دیماند (KW)", "بار صبح (KW)", "بار شب (KW)", "کاهش (KW)", "درصد کاهش (%)"];
    const tableRows = [];

    parsedData.forEach(customer => {
        const customerData = [
            customer.bodyNumber,
            customer.customerName,
            customer.billId,
            customer.contractDemand,
            customer.morningLoad,
            customer.eveningLoad,
            customer.reductionKW,
            customer.reductionPercent
        ];
        tableRows.push(customerData);
    });

    // اضافه کردن جدول
    doc.autoTable({
        head: [tableColumn],
        body: tableRows,
        startY: 60,
        theme: 'grid',
        styles: { font: 'Amiri', fontStyle: 'normal', halign: 'center' }, // استفاده از فونت فارسی
        headStyles: { fillColor: [22, 160, 133] },
        margin: { top: 50 },
        didDrawPage: function (data) {
            doc.text("گزارش تحلیل بار مشترکین", 40, 30); // عنوان صفحه
        }
    });

    // اضافه کردن نمودارها به PDF
    // این قسمت نیاز به پیچیدگی بیشتری دارد، زیرا Chart.js مستقیماً به PDF تبدیل نمی‌شود.
    // شما باید هر نمودار را به تصویر (Image) تبدیل کنید و سپس تصاویر را به PDF اضافه کنید.
    // این یک مثال کلی است و ممکن است نیاز به تنظیمات بیشتری داشته باشد:
    let yOffset = doc.autoTable.previous.finalY + 30; // شروع بعد از جدول

    for (const customer of parsedData) {
        const chartId = `chart-${customer.bodyNumber}`;
        const canvas = document.getElementById(chartId);
        if (canvas) {
            try {
                const imgData = canvas.toDataURL('image/png', 1.0); // تبدیل کانواس به تصویر PNG
                const imgWidth = 500; // عرض تصویر در PDF
                const imgHeight = (canvas.height * imgWidth) / canvas.width; // حفظ نسبت تصویر
                const margin = 40; // حاشیه
                const availableWidth = doc.internal.pageSize.getWidth() - 2 * margin;

                // اگر تصویر از صفحه خارج می‌شود، صفحه جدید اضافه کنید
                if (yOffset + imgHeight > doc.internal.pageSize.getHeight() - margin) {
                    doc.addPage();
                    yOffset = margin; // شروع از بالای صفحه جدید
                }

                // عنوان نمودار
                doc.setFontSize(12);
                doc.text(`نمودار بار مصرفی مشترک: ${customer.customerName} (شماره بدنه: ${customer.bodyNumber})`, margin, yOffset);
                yOffset += 20;

                // اضافه کردن تصویر نمودار
                doc.addImage(imgData, 'PNG', margin, yOffset, availableWidth, imgHeight);
                yOffset += imgHeight + 30; // فاصله برای نمودار بعدی
                console.log(`نمودار مشتری ${customer.bodyNumber} به PDF اضافه شد.`); // لاگ: نمودار به PDF
            } catch (error) {
                console.error(`خطا در تبدیل نمودار مشتری ${customer.bodyNumber} به تصویر برای PDF:`, error);
                Swal.fire('خطا', `مشکل در افزودن نمودار مشتری ${customer.customerName} به PDF.`, 'error');
            }
        }
    }

    doc.save("گزارش_تحلیل_بار.pdf");
    console.log("خروجی PDF ایجاد شد."); // لاگ: خروجی PDF
}

// رویدادها برای به‌روزرسانی لحظه‌ای مقادیر فیلتر
chkEvening.addEventListener('change', () => txtEvening.disabled = !chkEvening.checked);
chkReduction.addEventListener('change', () => txtReduction.disabled = !chkReduction.checked);

// تنظیمات اولیه هنگام بارگذاری صفحه
document.addEventListener('DOMContentLoaded', () => {
    showProgress(0); // مخفی کردن نوار پیشرفت در ابتدا
    sheetSelect.disabled = true; // غیرفعال کردن انتخاب شیت
    processBtn.disabled = true; // غیرفعال کردن دکمه پردازش
    exportExcelBtn.disabled = true;
    exportPdfBtn.disabled = true;
    noChartsMessage.style.display = 'block'; // نمایش پیام "نموداری وجود ندارد"
});