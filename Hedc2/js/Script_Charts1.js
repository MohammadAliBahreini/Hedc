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
    // دسترسی به عناصر نوار پیشرفت. این توابع به دلیل فراخوانی در DOMContentLoaded یا بعد از آن،
    // مطمئن هستند که عناصر DOM وجود دارند، اما برای اطمینان بیشتر، بررسی وجود انجام می‌شود.
    const progressBar = document.getElementById('progress-bar');
    const progressLabel = document.getElementById('progress-label');
    const progressContainer = document.getElementById('progress-container');

    if (progressBar && progressLabel && progressContainer) {
        progressBar.style.width = percent + '%';
        progressBar.setAttribute('aria-valuenow', percent);
        progressLabel.textContent = label;
        if (percent === 0 || percent === 100) {
            progressContainer.style.display = 'none'; // مخفی کردن نوار در شروع یا پایان
        } else {
            progressContainer.style.display = 'block'; // نمایش نوار در حین عملیات
        }
    }
}

/**
 * تابع برای از بین بردن (Destroy) نمودارهای Chart.js قبلی
 * این کار برای جلوگیری از انباشت نمودارها و مصرف حافظه اضافی ضروری است.
 */
function destroyCharts() {
    currentCharts.forEach(chart => {
        if (chart) {
            chart.destroy(); // متد destroy برای پاکسازی صحیح نمودارها
        }
    });
    currentCharts = []; // پاکسازی آرایه نگهداری نمودارها
    const chartsContainer = document.getElementById('chartsContainer');
    if (chartsContainer) {
        chartsContainer.innerHTML = ''; // پاک کردن محتوای HTML کانتینر نمودارها
    }
}

/**
 * تابع اصلی پردازش داده‌ها از فایل اکسل
 * این تابع پس از انتخاب شیت و کلیک دکمه "پردازش" فراخوانی می‌شود.
 */
async function processData() {
    // دسترسی به عناصر DOM که در این تابع استفاده می‌شوند
    const sheetSelect = document.getElementById('sheetSelect');
    const resultsTableBody = document.querySelector('#resultsTable tbody');
    const chartsContainer = document.getElementById('chartsContainer');
    const noChartsMessage = document.getElementById('noChartsMessage');
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


    // بررسی وجود ورک‌بوک
    if (!workbook) {
        Swal.fire('خطا', 'لطفاً ابتدا یک فایل اکسل انتخاب کنید.', 'error');
        return;
    }

    // انتخاب شیت جاری یا اولین شیت در صورت عدم انتخاب
    const selectedSheetName = sheetSelect.value || workbook.SheetNames[0];
    if (!selectedSheetName) {
        Swal.fire('خطا', 'شیتی برای پردازش یافت نشد.', 'error');
        return;
    }

    showProgress(25, 'در حال پردازش داده‌ها...');

    // پاکسازی نتایج و نمودارهای قبلی قبل از شروع پردازش جدید
    if (resultsTableBody) resultsTableBody.innerHTML = '';
    destroyCharts();
    if (chartsContainer) chartsContainer.innerHTML = '';
    if (noChartsMessage) noChartsMessage.style.display = 'none';

    // خواندن داده‌ها از شیت انتخاب‌شده
    const worksheet = workbook.Sheets[selectedSheetName];
    // sheet_to_json با {header: 1} یک آرایه از آرایه‌ها برمی‌گرداند که ردیف اول سربرگ است
    // raw: true برای دریافت مقادیر خام (بدون فرمت‌دهی اکسل)
    const jsonSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });

    const headers = jsonSheet[0]; // ردیف اول حاوی سربرگ‌ها است
    console.log("سربرگ‌ها از اکسل:", headers);
    const dataRows = jsonSheet.slice(1); // بقیه ردیف‌ها حاوی داده‌ها هستند
    console.log("ردیف‌های داده (dataRows):", dataRows);

    parsedData = []; // پاکسازی آرایه داده‌های پردازش‌شده برای شروع مجدد

    // محاسبه مرزهای زمانی (صبح و شب) بر اساس ورودی‌های کاربر
    const getMinutes = (h, m) => parseInt(h) * 60 + parseInt(m);
    const morningStart = getMinutes(morningStartHour.value, morningStartMinute.value);
    const morningEnd = getMinutes(morningEndHour.value, morningEndMinute.value);
    const eveningStart = getMinutes(eveningStartHour.value, eveningStartMinute.value);
    const eveningEnd = getMinutes(eveningEndHour.value, eveningEndMinute.value);

    // شناسایی ایندکس ستون‌های ضروری بر اساس نام سربرگ
    // *نکته مهم: این نام‌ها باید دقیقاً با سربرگ‌های فایل اکسل مطابقت داشته باشند.*
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

    // دریافت مقادیر فیلترها از UI
    const minEveningLoad = chkEvening && chkEvening.checked ? parseFloat(txtEvening.value) : -Infinity;
    const minReductionPercent = chkReduction && chkReduction.checked ? parseFloat(txtReduction.value) : -Infinity;

    showProgress(50, 'در حال تحلیل بارهای مشترکین...');

    // حلقه بر روی هر ردیف داده برای پردازش
    for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
        // از ردیف‌های خالی یا ردیف‌هایی که شماره بدنه ندارند بگذرید
        if (!row || row.length === 0 || row[bodyNumberColIndex] === undefined || row[bodyNumberColIndex] === null) {
            console.warn(`ردیف ${i + 2} به دلیل خالی بودن یا نداشتن شماره بدنه، نادیده گرفته شد.`);
            continue;
        }

        // جمع‌آوری اطلاعات اصلی مشترک
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
        const timeLabels = []; // برای ذخیره برچسب‌های زمانی (مانند "00:00") برای نمودار

        // شناسایی اولین ستون داده بار
        // ایندکس آخرین ستون اطلاعات ثابت + 1 (فرض بر این است که ستون‌های بار بلافاصله بعد از اینها شروع می‌شوند)
        const firstLoadColumnIndex = Math.max(
            bodyNumberColIndex, customerNameColIndex, billIdColIndex,
            addressColIndex, subscriptionNumberColIndex, demandColumnIndex
        ) + 1;

        // حلقه بر روی ستون‌های داده بار
        for (let j = firstLoadColumnIndex; j < headers.length; j++) {
            const header = headers[j];
            // بررسی کنید که آیا سربرگ شبیه یک زمان است (مثلاً "00:00", "00:15")
            const timeMatch = String(header).match(/^(\d{2}:\d{2}) to \d{2}:\d{2} \[KW\]$/);

        if (timeMatch && timeMatch[1]) { // اگر الگو مطابقت داشت و گروه اول (زمان شروع) وجود داشت
            const timeString = timeMatch[1]; // "00:00"
            const [h, m] = timeString.split(':').map(Number);
            const timeInMinutes = h * 60 + m;
            const loadValue = parseFloat(row[j]);

            if (!isNaN(loadValue)) {
                loadProfile.push({ timeInMinutes, load: loadValue });
                timeLabels.push(timeString); // استفاده از رشته "HH:MM" برای برچسب نمودار
            } else {
                console.warn(`مقدار بار نامعتبر در ردیف ${i + 2}, ستون ${header}: ${row[j]}`);
            }
        } else {
            // این خط برای دیباگینگ مفید است، نشان می‌دهد کدام سربرگ‌ها نادیده گرفته می‌شوند
            console.warn(`سربرگ "${header}" به عنوان ستون زمان بار معتبر شناسایی نشد.`);
        }
        }

        if (loadProfile.length === 0) {
            console.warn(`ردیف ${i + 2} (${customerInfo.bodyNumber}) به دلیل پروفایل بار خالی (عدم یافتن ستون‌های زمانی معتبر)، نادیده گرفته شد.`);
            continue; // اگر داده باری یافت نشد، از این مشتری بگذرید
        }

        // محاسبه بار صبح بر اساس نوع محاسبه (میانگین، حداکثر، حداقل)
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

        // محاسبه کاهش بار و درصد کاهش
        const reductionKW = morningLoad - eveningLoad;
        const reductionPercent = (morningLoad > 0) ? (reductionKW / morningLoad) * 100 : 0;

        // ساخت شیء نتیجه برای هر مشترک
        const customerResult = {
            ...customerInfo, // کپی کردن اطلاعات اصلی مشترک
            morningLoad: morningLoad.toFixed(2), // گرد کردن به 2 رقم اعشار
            eveningLoad: eveningLoad.toFixed(2),
            reductionKW: reductionKW.toFixed(2),
            reductionPercent: reductionPercent.toFixed(2),
            loadProfileData: loadProfile.map(item => item.load), // فقط مقادیر بار برای داده‌های نمودار
            timeLabels: timeLabels // برچسب‌های زمانی برای محور X نمودار
        };

        // اعمال فیلترها (اگر تیک خورده باشند و مقدار معتبر داشته باشند)
        const passesEveningFilter = !chkEvening || !chkEvening.checked || (parseFloat(customerResult.eveningLoad) >= minEveningLoad);
        const passesReductionFilter = !chkReduction || !chkReduction.checked || (parseFloat(customerResult.reductionPercent) >= minReductionPercent);

        if (passesEveningFilter && passesReductionFilter) {
            parsedData.push(customerResult); // اضافه کردن مشترک به لیست نهایی در صورت عبور از فیلترها
            console.log("مشتری با موفقیت اضافه شد:", customerResult);
        } else {
            console.log(`مشتری ${customerInfo.bodyNumber} به دلیل عدم تطابق با فیلترها اضافه نشد.`, {
                passesEveningFilter,
                passesReductionFilter,
                eveningLoad: customerResult.eveningLoad,
                minEveningLoad,
                reductionPercent: customerResult.reductionPercent,
                minReductionPercent
            });
        }
    } // پایان حلقه for (let i = 0; i < dataRows.length; i++)

    console.log("آرایه نهایی parsedData پس از پردازش:", parsedData);

    showProgress(75, 'در حال نمایش نتایج...');
    displayResults(); // نمایش نتایج در جدول
    drawCharts(); // رسم نمودارها
    showProgress(100, 'پردازش کامل شد.');

    // فعال کردن دکمه‌های خروجی گرفتن
    if (exportExcelBtn) exportExcelBtn.disabled = false;
    if (exportPdfBtn) exportPdfBtn.disabled = false;
}

/**
 * تابع برای نمایش نتایج پردازش شده در جدول HTML
 */
function displayResults() {
    const resultsTableBody = document.querySelector('#resultsTable tbody');

    if (resultsTableBody) {
        resultsTableBody.innerHTML = ''; // پاکسازی جدول قبل از افزودن نتایج جدید

        if (parsedData.length === 0) {
            const row = resultsTableBody.insertRow();
            const cell = row.insertCell();
            cell.colSpan = 11; // تعداد ستون‌ها (شامل ردیف و نمودار)
            cell.textContent = 'هیچ داده‌ای بر اساس فیلترهای اعمال شده یافت نشد.';
            cell.style.textAlign = 'center';
            console.log("هیچ داده‌ای برای نمایش در جدول وجود ندارد.");
            return;
        }

        parsedData.forEach(customer => {
            const row = resultsTableBody.insertRow();
            row.insertCell().textContent = customer.rowNum; // شماره ردیف در اکسل
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
            // با کلیک بر روی دکمه، به نمودار مربوطه اسکرول می‌کند
            viewChartBtn.onclick = () => scrollToChart(customer.bodyNumber);
            chartCell.appendChild(viewChartBtn);
        });
        console.log(`تعداد ${parsedData.length} مشتری در جدول نمایش داده شد.`);
    }
}

/**
 * تابع برای رسم نمودارهای بار مصرفی با Chart.js
 */
function drawCharts() {
    destroyCharts(); // ابتدا هر نمودار موجود را پاک کنید

    const noChartsMessage = document.getElementById('noChartsMessage');
    const chartsContainer = document.getElementById('chartsContainer');

    if (parsedData.length === 0) {
        if (noChartsMessage) noChartsMessage.style.display = 'block'; // نمایش پیام "نموداری وجود ندارد"
        console.log("هیچ داده‌ای برای رسم نمودار وجود ندارد.");
        return;
    }

    if (noChartsMessage) noChartsMessage.style.display = 'none'; // مخفی کردن پیام

    parsedData.forEach(customer => {
        const chartId = `chart-${customer.bodyNumber}`; // ID منحصر به فرد برای هر نمودار
        const chartDiv = document.createElement('div');
        chartDiv.className = 'chart-container';
        chartDiv.id = `chart-div-${customer.bodyNumber}`; // ID برای اسکرول کردن به آن
        chartDiv.innerHTML = `
            <h3>نمودار بار مصرفی مشترک: ${customer.customerName} (شماره بدنه: ${customer.bodyNumber})</h3>
            <canvas id="${chartId}"></canvas>
            <hr>
        `;
        if (chartsContainer) chartsContainer.appendChild(chartDiv);

        const ctx = document.getElementById(chartId);
        if (ctx) { // مطمئن شوید عنصر canvas پیدا شده است
            const newChart = new Chart(ctx.getContext('2d'), {
                type: 'line', // نوع نمودار خطی
                data: {
                    labels: customer.timeLabels, // برچسب‌های زمان برای محور X
                    datasets: [{
                        label: 'بار مصرفی (KW)',
                        data: customer.loadProfileData, // داده‌های بار مصرفی
                        borderColor: 'rgb(75, 192, 192)', // رنگ خط نمودار
                        tension: 0.1, // خمیدگی خط
                        fill: false // عدم پر کردن ناحیه زیر خط
                    }]
                },
                options: {
                    responsive: true, // واکنش‌گرا (responsive) بودن نمودار
                    maintainAspectRatio: false, // عدم حفظ نسبت ابعاد برای کنترل بهتر اندازه
                    scales: {
                        x: {
                            title: {
                                display: true,
                                text: 'زمان' // عنوان محور X
                            }
                        },
                        y: {
                            title: {
                                display: true,
                                text: 'بار (KW)' // عنوان محور Y
                            },
                            beginAtZero: true // شروع محور Y از صفر
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
            currentCharts.push(newChart); // اضافه کردن نمونه نمودار به آرایه برای مدیریت بعدی
            console.log(`نمودار برای مشتری ${customer.bodyNumber} رسم شد.`);
        } else {
            console.warn(`عنصر canvas برای نمودار مشتری ${customer.bodyNumber} یافت نشد.`);
        }
    });
    console.log(`تعداد ${currentCharts.length} نمودار رسم شد.`);
}

/**
 * تابع برای اسکرول کردن به نمودار خاص در صفحه
 * @param {string} bodyNumber - شماره بدنه مشترک که نمودارش باید نمایش داده شود.
 */
function scrollToChart(bodyNumber) {
    const chartDiv = document.getElementById(`chart-div-${bodyNumber}`);
    if (chartDiv) {
        // اسکرول به عنصر با رفتار نرم (smooth) و قرار گرفتن در بالای صفحه (start)
        chartDiv.scrollIntoView({ behavior: 'smooth', block: 'start' });
        console.log(`به نمودار مشتری ${bodyNumber} اسکرول شد.`);
    } else {
        console.warn(`نمودار مشتری ${bodyNumber} یافت نشد.`);
    }
}

/**
 * تابع برای خروجی گرفتن نتایج پردازش شده به فرمت اکسل
 */
function exportToExcel() {
    if (parsedData.length === 0) {
        Swal.fire('هشدار', 'داده‌ای برای خروجی اکسل وجود ندارد.', 'warning');
        return;
    }

    // تعریف سربرگ‌های فایل اکسل خروجی
    const wsData = [
        ["ردیف", "شماره بدنه", "نام مشترک", "شناسه قبض", "آدرس مشترک", "شماره اشتراک", "دیماند قراردادی (KW)", "بار صبح (KW)", "بار شب (KW)", "کاهش بار (KW)", "درصد کاهش بار (%)"]
    ];
    // افزودن داده‌های پردازش شده به آرایه داده‌های اکسل
    parsedData.forEach(customer => {
        wsData.push([
            customer.rowNum,
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

    const ws = XLSX.utils.aoa_to_sheet(wsData); // تبدیل آرایه آرایه‌ها به شیء شیت
    const wb = XLSX.utils.book_new(); // ایجاد یک ورک‌بوک جدید
    XLSX.utils.book_append_sheet(wb, ws, "نتایج"); // اضافه کردن شیت به ورک‌بوک
    XLSX.writeFile(wb, "نتایج_تحلیل_بار.xlsx"); // ذخیره فایل اکسل
    console.log("خروجی اکسل ایجاد شد.");
}

/**
 * تابع برای خروجی گرفتن نتایج و نمودارها به فرمت PDF
 * از کتابخانه‌های jsPDF و jspdf-autotable استفاده می‌کند.
 * برای نمودارها، canvas نمودار را به تصویر تبدیل کرده و به PDF اضافه می‌کند.
 */
async function exportToPdf() {
    if (parsedData.length === 0) {
        Swal.fire('هشدار', 'داده‌ای برای خروجی PDF وجود ندارد.', 'warning');
        return;
    }

    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'pt', 'a4'); // ایجاد سند PDF جدید (پرتره، واحد pt، سایز A4)

    // اضافه کردن فونت فارسی (Amiri) برای پشتیبانی از متن فارسی در PDF
    // مطمئن شوید فایل 'Amiri-Regular.ttf' در مسیری قابل دسترس برای وب‌سرور شما قرار دارد.
    // و این فونت توسط jsPDF در هنگام کامپایل/باندل کردن جاسازی شده باشد (اگر از CDN استفاده نمی‌کنید)
    doc.addFont('fonts/Amiri-Regular.ttf', 'Amiri', 'normal');
    doc.setFont('Amiri');

    // تعریف ستون‌های جدول برای PDF
    const tableColumn = ["ردیف", "شماره بدنه", "نام مشترک", "شناسه قبض", "دیماند (KW)", "بار صبح (KW)", "بار شب (KW)", "کاهش (KW)", "درصد کاهش (%)"];
    const tableRows = [];

    // پر کردن ردیف‌های جدول با داده‌های پردازش شده
    parsedData.forEach(customer => {
        const customerData = [
            customer.rowNum,
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

    // اضافه کردن جدول به PDF با jsPDF-autotable
    doc.autoTable({
        head: [tableColumn], // سربرگ جدول
        body: tableRows, // ردیف‌های داده
        startY: 60, // شروع جدول از ارتفاع 60pt
        theme: 'grid', // تم جدول
        styles: { font: 'Amiri', fontStyle: 'normal', halign: 'center', cellPadding: 5, fontSize: 8 }, // استایل‌های سلول
        headStyles: { fillColor: [22, 160, 133], fontSize: 9 }, // استایل‌های سربرگ
        margin: { top: 50, right: 30, left: 30 }, // حاشیه‌های صفحه
        didDrawPage: function (data) {
            // اضافه کردن عنوان به هر صفحه
            doc.setFontSize(14);
            doc.text("گزارش تحلیل بار مشترکین", doc.internal.pageSize.getWidth() / 2, 30, { align: "center" });
        }
    });

    let yOffset = doc.autoTable.previous.finalY + 30; // نقطه شروع برای نمودارها (بعد از جدول)

    // اضافه کردن نمودارها به PDF
    for (const customer of parsedData) {
        const chartId = `chart-${customer.bodyNumber}`;
        const canvas = document.getElementById(chartId);

        if (canvas) {
            try {
                // یک تأخیر کوچک برای اطمینان از اینکه Chart.js نمودار را کاملاً رندر کرده است
                await new Promise(resolve => setTimeout(resolve, 50));

                const imgData = canvas.toDataURL('image/png', 1.0); // تبدیل canvas به تصویر PNG
                const imgWidth = 500; // عرض تصویر در PDF
                const imgHeight = (canvas.height * imgWidth) / canvas.width; // حفظ نسبت ابعاد
                const margin = 40;
                const availableWidth = doc.internal.pageSize.getWidth() - 2 * margin;

                // بررسی کنید که آیا نمودار در صفحه فعلی جا می‌شود یا نه. اگر نه، صفحه جدید اضافه کنید.
                if (yOffset + imgHeight + 50 > doc.internal.pageSize.getHeight() - margin) {
                    doc.addPage();
                    yOffset = margin; // شروع از بالای صفحه جدید
                }

                // عنوان نمودار در PDF
                doc.setFontSize(12);
                doc.text(`نمودار بار مصرفی مشترک: ${customer.customerName} (شماره بدنه: ${customer.bodyNumber})`, margin, yOffset);
                yOffset += 20; // فاصله بین عنوان و نمودار

                doc.addImage(imgData, 'PNG', margin, yOffset, availableWidth, imgHeight); // اضافه کردن تصویر نمودار به PDF
                yOffset += imgHeight + 30; // فاصله برای نمودار بعدی
                console.log(`نمودار مشتری ${customer.bodyNumber} به PDF اضافه شد.`);
            } catch (error) {
                console.error(`خطا در تبدیل نمودار مشتری ${customer.bodyNumber} به تصویر برای PDF:`, error);
                Swal.fire('خطا', `مشکل در افزودن نمودار مشتری ${customer.customerName} به PDF.`, 'error');
            }
        } else {
            console.warn(`عنصر Canvas برای نمودار مشتری ${customer.bodyNumber} در DOM یافت نشد.`);
        }
    }

    doc.save("گزارش_تحلیل_بار.pdf"); // ذخیره فایل PDF
    console.log("خروجی PDF ایجاد شد.");
}

// ====================================================================================================
// اجرای کدهای DOM-ready (پس از بارگذاری کامل DOM)
// تمام کدهای مربوط به دسترسی به عناصر HTML و اضافه کردن Event Listenerها باید اینجا باشند.
// ====================================================================================================
document.addEventListener('DOMContentLoaded', () => {
    // دسترسی به عناصر DOM
    // این متغیرها به صورت 'const' (ثابت) تعریف می‌شوند زیرا به عنصر خاصی از DOM اشاره می‌کنند
    // و ارجاع آن‌ها تغییر نخواهد کرد.
    const fileInput = document.getElementById('excelFile'); // ID جدید
    const sheetSelect = document.getElementById('sheetSelect');
    const processBtn = document.getElementById('processDataBtn'); // ID جدید
    const resultsTableBody = document.querySelector('#resultsTable tbody'); // انتخاب tbody
    const progressContainer = document.getElementById('progress-container'); // ID جدید
    const progressBar = document.getElementById('progress-bar'); // ID جدید
    const progressLabel = document.getElementById('progress-label'); // ID جدید
    const chartsContainer = document.getElementById('chartsContainer');
    const noChartsMessage = document.getElementById('noChartsMessage');
    const exportExcelBtn = document.getElementById('exportExcelBtn');
    const exportPdfBtn = document.getElementById('exportPdfBtn');
    const resetAppBtn = document.getElementById('resetAppBtn'); // دکمه جدید برای شروع مجدد
    const fileNameDisplay = document.getElementById('fileNameDisplay'); // برای نمایش نام فایل انتخاب شده

    // فیلترها و تنظیمات محاسبه (عناصر DOM مربوطه)
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

    // تنظیمات اولیه هنگام بارگذاری صفحه
    showProgress(0); // مخفی کردن نوار پیشرفت در ابتدا

    // غیرفعال کردن اولیه دکمه‌ها و دراپ‌داون‌ها
    if (sheetSelect) sheetSelect.disabled = true;
    if (processBtn) processBtn.disabled = true;
    if (exportExcelBtn) exportExcelBtn.disabled = true;
    if (exportPdfBtn) exportPdfBtn.disabled = true;
    if (noChartsMessage) noChartsMessage.style.display = 'block'; // نمایش پیام "نموداری وجود ندارد"

    // ====================================================================================================
    // Event Listeners (گوش دادن به رویدادها)
    // ====================================================================================================

    // گوش دادن به رویداد تغییر (change) فایل ورودی
    if (fileInput) { // بررسی وجود عنصر
        fileInput.addEventListener('change', async (event) => {
            const file = event.target.files[0];
            if (!file) {
                // اگر فایلی انتخاب نشد، حالت اولیه را برگردانید
                if (fileNameDisplay) fileNameDisplay.textContent = 'فایل انتخاب نشده...';
                if (sheetSelect) sheetSelect.innerHTML = '<option value="">- Sheet1 -</option>';
                if (sheetSelect) sheetSelect.disabled = true;
                if (processBtn) processBtn.disabled = true;
                showProgress(0);
                return;
            }

            if (fileNameDisplay) fileNameDisplay.textContent = file.name; // نمایش نام فایل انتخاب شده
            showProgress(10, 'در حال خواندن فایل...');

            const reader = new FileReader(); // ایجاد شیء FileReader برای خواندن فایل
            reader.onload = async (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    workbook = XLSX.read(data, { type: 'array' }); // خواندن ورک‌بوک اکسل
                    console.log("ورک‌بوک با موفقیت بارگذاری شد:", workbook);

                    // پر کردن دراپ‌داون شیت‌ها
                    if (sheetSelect) {
                        sheetSelect.innerHTML = ''; // پاکسازی گزینه‌های قبلی
                        workbook.SheetNames.forEach(sheetName => {
                            const option = document.createElement('option');
                            option.value = sheetName;
                            option.textContent = sheetName;
                            sheetSelect.appendChild(option);
                        });
                        sheetSelect.disabled = false; // فعال کردن دراپ‌داون شیت
                    }
                    if (processBtn) processBtn.disabled = false; // فعال کردن دکمه پردازش
                    showProgress(100, 'فایل آماده پردازش است.');
                } catch (error) {
                    console.error("خطا در بارگذاری فایل اکسل:", error);
                    Swal.fire('خطا', 'فایل اکسل نامعتبر است یا در خواندن آن مشکلی پیش آمده.', 'error');
                    showProgress(0);
                    // ریست کردن UI در صورت خطا
                    if (fileNameDisplay) fileNameDisplay.textContent = 'فایل انتخاب نشده...';
                    if (sheetSelect) sheetSelect.innerHTML = '<option value="">- Sheet1 -</option>';
                    if (sheetSelect) sheetSelect.disabled = true;
                    if (processBtn) processBtn.disabled = true;
                }
            };
            reader.readAsArrayBuffer(file); // شروع خواندن فایل به عنوان ArrayBuffer
        });
    }

    // گوش دادن به رویداد کلیک دکمه پردازش
    if (processBtn) processBtn.addEventListener('click', processData);
    // گوش دادن به رویداد کلیک دکمه‌های خروجی
    if (exportExcelBtn) exportExcelBtn.addEventListener('click', exportToExcel);
    if (exportPdfBtn) exportPdfBtn.addEventListener('click', exportToPdf);

    // رویدادها برای به‌روزرسانی وضعیت فیلترها (فعال/غیرفعال کردن فیلدهای ورودی)
    if (chkEvening) chkEvening.addEventListener('change', () => {
        if (txtEvening) txtEvening.disabled = !chkEvening.checked;
    });
    if (chkReduction) chkReduction.addEventListener('change', () => {
        if (txtReduction) txtReduction.disabled = !chkReduction.checked;
    });

    // گوش دادن به رویداد کلیک دکمه "شروع مجدد"
    if (resetAppBtn) {
        resetAppBtn.addEventListener('click', () => {
            // ریست کردن تمام متغیرهای سراسری
            workbook = null;
            parsedData = [];
            currentCharts = [];

            // ریست کردن عناصر UI به حالت اولیه
            if (fileInput) {
                fileInput.value = ''; // پاک کردن فایل انتخاب شده از ورودی
                if (fileNameDisplay) fileNameDisplay.textContent = 'فایل انتخاب نشده...';
            }
            if (sheetSelect) {
                sheetSelect.innerHTML = '<option value="">- Sheet1 -</option>';
                sheetSelect.disabled = true;
            }
            if (processBtn) processBtn.disabled = true;
            if (exportExcelBtn) exportExcelBtn.disabled = true;
            if (exportPdfBtn) exportPdfBtn.disabled = true;
            if (resultsTableBody) resultsTableBody.innerHTML = ''; // پاکسازی جدول نتایج
            destroyCharts(); // پاکسازی نمودارها
            if (noChartsMessage) noChartsMessage.style.display = 'block'; // نمایش پیام "نموداری وجود ندارد"

            // ریست کردن فیلترها به حالت پیش‌فرض (اختیاری)
            if (chkEvening) chkEvening.checked = false;
            if (txtEvening) txtEvening.disabled = true;
            if (chkReduction) chkReduction.checked = false;
            if (txtReduction) txtReduction.disabled = true;

            showProgress(0, 'منتظر انتخاب فایل...'); // ریست نوار پیشرفت
            Swal.fire('با موفقیت', 'برنامه به حالت اولیه بازگردانده شد.', 'success');
        });
    }

}); // پایان DOMContentLoaded