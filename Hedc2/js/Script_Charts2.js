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
            progressContainer.style.display = 'none';
        } else {
            progressContainer.style.display = 'block';
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
            chart.destroy();
        }
    });
    currentCharts = [];
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
    if (chartsContainer) chartsContainer.innerHTML = '';
    if (noChartsMessage) noChartsMessage.style.display = 'none';

    const worksheet = workbook.Sheets[selectedSheetName];
    const jsonSheet = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: true });

    const headers = jsonSheet[0];
    console.log("سربرگ‌ها از اکسل:", headers);
    const dataRows = jsonSheet.slice(1);
    console.log("ردیف‌های داده (dataRows):", dataRows);

    parsedData = []; // پاکسازی آرایه داده‌های پردازش‌شده برای شروع مجدد

    const getMinutes = (h, m) => parseInt(h) * 60 + parseInt(m);
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
    const minReductionPercent = chkReduction && chkReduction.checked ? parseFloat(txtReduction.value) : -Infinity;

    showProgress(50, 'در حال تحلیل بارهای مشترکین...');

    for (let i = 0; i < dataRows.length; i++) {
        const row = dataRows[i];
        if (!row || row.length === 0 || row[bodyNumberColIndex] === undefined || row[bodyNumberColIndex] === null) {
            console.warn(`ردیف ${i + 2} به دلیل خالی بودن یا نداشتن شماره بدنه، نادیده گرفته شد.`);
            continue;
        }

        const customerInfo = {
            id: `customer-${customerNameColIndex}-${i}`, // یک ID منحصر به فرد برای هر مشتری
            rowNum: i + 2,
            bodyNumber: row[bodyNumberColIndex],
            customerName: row[customerNameColIndex],
            billId: row[billIdColIndex],
            address: row[addressColIndex],
            subscriptionNumber: row[subscriptionNumberColIndex],
            contractDemand: parseFloat(row[demandColumnIndex]) || 0
        };

        const loadProfile = [];
        const timeLabels = [];

        const firstLoadColumnIndex = Math.max(
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
                    timeLabels.push(timeString);
                } else {
                    console.warn(`مقدار بار نامعتبر در ردیف ${i + 2}, ستون ${header}: ${row[j]}`);
                }
            } else {
                console.warn(`سربرگ "${header}" به عنوان ستون زمان بار معتبر شناسایی نشد.`);
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
        const passesReductionFilter = !chkReduction || !chkReduction.checked || (parseFloat(customerResult.reductionPercent) >= minReductionPercent);

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
                minReductionPercent
            });
        }
    }

    console.log("آرایه نهایی parsedData پس از پردازش:", parsedData);

    showProgress(75, 'در حال نمایش نتایج...');
    displayResults();
    drawCharts();
    showProgress(100, 'پردازش کامل شد.');

    if (exportExcelBtn) exportExcelBtn.disabled = false;
    if (exportPdfBtn) exportPdfBtn.disabled = false;
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
            cell.colSpan = 12; // افزایش تعداد ستون‌ها (شامل دکمه حذف)
            cell.textContent = 'هیچ داده‌ای بر اساس فیلترهای اعمال شده یافت نشد.';
            cell.style.textAlign = 'center';
            console.log("هیچ داده‌ای برای نمایش در جدول وجود ندارد.");
            return;
        }

        parsedData.forEach((customer, index) => {
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
            deleteBtn.onclick = () => deleteCustomerRow(customer.id); // فراخوانی تابع حذف با ID مشترک
            deleteCell.appendChild(deleteBtn);
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
        if (noChartsMessage) noChartsMessage.style.display = 'block';
        console.log("هیچ داده‌ای برای رسم نمودار وجود ندارد.");
        return;
    }

    if (noChartsMessage) noChartsMessage.style.display = 'none';

    parsedData.forEach(customer => {
        const chartId = `chart-${customer.bodyNumber}`; // ID منحصر به فرد برای هر نمودار
        const chartDiv = document.createElement('div');
        chartDiv.className = 'chart-container my-3 p-3 border rounded shadow-sm bg-light'; // استایل‌های جدید برای هر کانتینر
        chartDiv.id = `chart-div-${customer.bodyNumber}`;
        chartDiv.innerHTML = `
            <h3 class="text-primary">نمودار بار مصرفی مشترک: ${customer.customerName} (شماره بدنه: ${customer.bodyNumber})</h3>
            <p class="text-muted">آدرس: ${customer.address}</p>
            <p class="text-info">کاهش بار: ${customer.reductionKW} KW &nbsp; | &nbsp; درصد کاهش: ${customer.reductionPercent}%</p>
            <div style="width: 100%; max-width: 800px; height: 400px; margin: auto;">
                <canvas id="${chartId}"></canvas>
            </div>
            <hr>
        `;
        if (chartsContainer) chartsContainer.appendChild(chartDiv);

        const ctx = document.getElementById(chartId);
        if (ctx) {
            const newChart = new Chart(ctx.getContext('2d'), {
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
                        },
                        // پلاگین annotation برای رسم خط دیماند
                        annotation: {
                            annotations: {
                                line1: {
                                    type: 'line',
                                    yMin: customer.contractDemand,
                                    yMax: customer.contractDemand,
                                    borderColor: 'rgb(255, 99, 132)', // رنگ قرمز
                                    borderWidth: 2,
                                    label: {
                                        content: `دیماند قراردادی: ${customer.contractDemand} KW`,
                                        enabled: true,
                                        position: 'start',
                                        backgroundColor: 'rgba(255, 99, 132, 0.8)',
                                        font: {
                                            size: 10
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            });
            currentCharts.push(newChart);
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
        chartDiv.scrollIntoView({ behavior: 'smooth', block: 'start' });
        console.log(`به نمودار مشتری ${bodyNumber} اسکرول شد.`);
    } else {
        console.warn(`نمودار مشتری ${bodyNumber} یافت نشد.`);
    }
}

/**
 * تابع برای حذف یک ردیف مشترک از جدول و نمودار مربوطه
 * @param {string} customerId - ID منحصر به فرد مشترک (تعریف شده در parsedData.id)
 */
function deleteCustomerRow(customerId) {
    Swal.fire({
        title: 'آیا مطمئن هستید؟',
        text: "این ردیف و نمودار مربوطه حذف خواهد شد!",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#d33',
        cancelButtonColor: '#3085d6',
        confirmButtonText: 'بله، حذف کن!',
        cancelButtonText: 'لغو'
    }).then((result) => {
        if (result.isConfirmed) {
            // حذف از آرایه parsedData
            const initialLength = parsedData.length;
            parsedData = parsedData.filter(customer => customer.id !== customerId);
            if (parsedData.length === initialLength) {
                console.warn(`مشتری با ID ${customerId} در parsedData یافت نشد.`);
                Swal.fire('خطا', 'مشکلی در حذف داده رخ داد. لطفاً صفحه را رفرش کنید.', 'error');
                return;
            }

            // حذف ردیف از جدول HTML
            const rowToRemove = document.querySelector(`#resultsTable tbody tr[data-customer-id="${customerId}"]`);
            if (rowToRemove) {
                rowToRemove.remove();
                console.log(`ردیف جدول برای مشتری ${customerId} حذف شد.`);
            } else {
                console.warn(`ردیف جدول برای مشتری ${customerId} یافت نشد.`);
            }

            // حذف نمودار از صفحه و آرایه currentCharts
            const chartDivToRemove = document.getElementById(`chart-div-${parsedData.find(c => c.id === customerId)?.bodyNumber || customerId.split('-')[1]}`); // پیدا کردن div بر اساس ID
            if (chartDivToRemove) {
                const canvasId = `chart-${parsedData.find(c => c.id === customerId)?.bodyNumber || customerId.split('-')[1]}`;
                const chartInstance = currentCharts.find(chart => chart.canvas.id === canvasId);
                if (chartInstance) {
                    chartInstance.destroy();
                    currentCharts = currentCharts.filter(chart => chart.canvas.id !== canvasId);
                    console.log(`نمودار برای مشتری ${customerId} حذف شد.`);
                }
                chartDivToRemove.remove();
            } else {
                // اگر مستقیم با bodyNumber ID ساخته شده باشد و customerId شامل آن نباشد، تلاش دوباره
                const bodyNumberFromId = customerId.split('-')[2]; // فرض کنید ID فرمتی شبیه 'customer-name-bodyNumber' دارد
                const fallbackChartDiv = document.getElementById(`chart-div-${bodyNumberFromId}`);
                if (fallbackChartDiv) {
                    const canvasId = `chart-${bodyNumberFromId}`;
                    const chartInstance = currentCharts.find(chart => chart.canvas.id === canvasId);
                    if (chartInstance) {
                        chartInstance.destroy();
                        currentCharts = currentCharts.filter(chart => chart.canvas.id !== canvasId);
                        console.log(`نمودار برای مشتری ${bodyNumberFromId} حذف شد (fallback).`);
                    }
                    fallbackChartDiv.remove();
                } else {
                    console.warn(`نمودار برای مشتری ${customerId} یا ${bodyNumberFromId} در DOM یافت نشد.`);
                }
            }


            // اگر هیچ داده‌ای باقی نمانده، پیام "هیچ داده‌ای نیست" را نمایش دهید
            if (parsedData.length === 0) {
                const resultsTableBody = document.querySelector('#resultsTable tbody');
                if (resultsTableBody) {
                    resultsTableBody.innerHTML = '';
                    const row = resultsTableBody.insertRow();
                    const cell = row.insertCell();
                    cell.colSpan = 12;
                    cell.textContent = 'هیچ داده‌ای بر اساس فیلترهای اعمال شده یافت نشد.';
                    cell.style.textAlign = 'center';
                }
                const noChartsMessage = document.getElementById('noChartsMessage');
                if (noChartsMessage) noChartsMessage.style.display = 'block';
                const exportExcelBtn = document.getElementById('exportExcelBtn');
                const exportPdfBtn = document.getElementById('exportPdfBtn');
                if (exportExcelBtn) exportExcelBtn.disabled = true;
                if (exportPdfBtn) exportPdfBtn.disabled = true;
            }

            Swal.fire('حذف شد!', 'ردیف و نمودار با موفقیت حذف شدند.', 'success');
        }
    });
}

/**
 * تابع برای خروجی گرفتن نتایج پردازش شده به فرمت اکسل
 */
function exportToExcel() {
    if (parsedData.length === 0) {
        Swal.fire('هشدار', 'داده‌ای برای خروجی اکسل وجود ندارد.', 'warning');
        return;
    }

    const wsData = [
        ["ردیف", "شماره بدنه", "نام مشترک", "شناسه قبض", "آدرس مشترک", "شماره اشتراک", "دیماند قراردادی (KW)", "بار صبح (KW)", "بار شب (KW)", "کاهش بار (KW)", "درصد کاهش بار (%)"]
    ];
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

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "نتایج");
    XLSX.writeFile(wb, "نتایج_تحلیل_بار.xlsx");
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
    const doc = new jsPDF('p', 'pt', 'a4');

    // اضافه کردن فونت فارسی (Amiri)
    doc.addFont('Amiri-Regular.ttf', 'Amiri', 'normal'); // مطمئن شوید این فایل در مسیر صحیح قرار دارد
    doc.setFont('Amiri');

    const tableColumn = ["ردیف", "شماره بدنه", "نام مشترک", "شناسه قبض", "آدرس", "دیماند (KW)", "بار صبح (KW)", "بار شب (KW)", "کاهش (KW)", "درصد کاهش (%)"];
    const tableRows = [];

    parsedData.forEach(customer => {
        const customerData = [
            customer.rowNum,
            customer.bodyNumber,
            customer.customerName,
            customer.billId,
            customer.address,
            customer.contractDemand,
            customer.morningLoad,
            customer.eveningLoad,
            customer.reductionKW,
            customer.reductionPercent
        ];
        tableRows.push(customerData);
    });

    doc.autoTable({
        head: [tableColumn],
        body: tableRows,
        startY: 60,
        theme: 'grid',
        styles: { font: 'Amiri', fontStyle: 'normal', halign: 'center', cellPadding: 5, fontSize: 8 },
        headStyles: { fillColor: [22, 160, 133], fontSize: 9 },
        margin: { top: 50, right: 30, left: 30 },
        didDrawPage: function (data) {
            doc.setFontSize(14);
            doc.text("گزارش تحلیل بار مشترکین", doc.internal.pageSize.getWidth() / 2, 30, { align: "center" });
        }
    });

    let yOffset = doc.autoTable.previous.finalY + 30;

    for (const customer of parsedData) {
        const chartId = `chart-${customer.bodyNumber}`;
        const canvas = document.getElementById(chartId);

        if (canvas) {
            try {
                // اطمینان از رندر کامل نمودار
                await new Promise(resolve => setTimeout(resolve, 100));

                const imgData = canvas.toDataURL('image/png', 1.0);
                const imgWidth = 500;
                const imgHeight = (canvas.height * imgWidth) / canvas.width;
                const margin = 40;
                const availableWidth = doc.internal.pageSize.getWidth() - 2 * margin;

                if (yOffset + imgHeight + 80 > doc.internal.pageSize.getHeight() - margin) { // 80 برای عنوان و آدرس
                    doc.addPage();
                    yOffset = margin;
                }

                doc.setFontSize(12);
                doc.text(`نمودار بار مصرفی مشترک: ${customer.customerName} (شماره بدنه: ${customer.bodyNumber})`, margin, yOffset);
                yOffset += 15;
                doc.setFontSize(10);
                doc.text(`آدرس: ${customer.address}`, margin, yOffset);
                yOffset += 15;
                doc.text(`کاهش بار: ${customer.reductionKW} KW | درصد کاهش: ${customer.reductionPercent}%`, margin, yOffset);
                yOffset += 20;

                doc.addImage(imgData, 'PNG', margin, yOffset, availableWidth, imgHeight);
                yOffset += imgHeight + 30;
                console.log(`نمودار مشتری ${customer.bodyNumber} به PDF اضافه شد.`);
            } catch (error) {
                console.error(`خطا در تبدیل نمودار مشتری ${customer.bodyNumber} به تصویر برای PDF:`, error);
                Swal.fire('خطا', `مشکل در افزودن نمودار مشتری ${customer.customerName} به PDF.`, 'error');
            }
        } else {
            console.warn(`عنصر Canvas برای نمودار مشتری ${customer.bodyNumber} در DOM یافت نشد.`);
        }
    }

    doc.save("گزارش_تحلیل_بار.pdf");
    console.log("خروجی PDF ایجاد شد.");
}


// ====================================================================================================
// اجرای کدهای DOM-ready (پس از بارگذاری کامل DOM)
// تمام کدهای مربوط به دسترسی به عناصر HTML و اضافه کردن Event Listenerها باید اینجا باشند.
// ====================================================================================================
document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excelFile');
    const sheetSelect = document.getElementById('sheetSelect');
    const processBtn = document.getElementById('processDataBtn');
    const resultsTableBody = document.querySelector('#resultsTable tbody');
    const progressContainer = document.getElementById('progress-container');
    const progressBar = document.getElementById('progress-bar');
    const progressLabel = document.getElementById('progress-label');
    const chartsContainer = document.getElementById('chartsContainer');
    const noChartsMessage = document.getElementById('noChartsMessage');
    const exportExcelBtn = document.getElementById('exportExcelBtn');
    const exportPdfBtn = document.getElementById('exportPdfBtn');
    const resetAppBtn = document.getElementById('resetAppBtn');
    const fileNameDisplay = document.getElementById('fileNameDisplay');

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

    showProgress(0);

    if (sheetSelect) sheetSelect.disabled = true;
    if (processBtn) processBtn.disabled = true;
    if (exportExcelBtn) exportExcelBtn.disabled = true;
    if (exportPdfBtn) exportPdfBtn.disabled = true;
    if (noChartsMessage) noChartsMessage.style.display = 'block';

    if (fileInput) {
        fileInput.addEventListener('change', async (event) => {
            const file = event.target.files[0];
            if (!file) {
                if (fileNameDisplay) fileNameDisplay.textContent = 'فایل انتخاب نشده...';
                if (sheetSelect) sheetSelect.innerHTML = '<option value="">- Sheet1 -</option>';
                if (sheetSelect) sheetSelect.disabled = true;
                if (processBtn) processBtn.disabled = true;
                showProgress(0);
                return;
            }

            if (fileNameDisplay) fileNameDisplay.textContent = file.name;
            showProgress(10, 'در حال خواندن فایل...');

            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    workbook = XLSX.read(data, { type: 'array' });
                    console.log("ورک‌بوک با موفقیت بارگذاری شد:", workbook);

                    if (sheetSelect) {
                        sheetSelect.innerHTML = '';
                        workbook.SheetNames.forEach(sheetName => {
                            const option = document.createElement('option');
                            option.value = sheetName;
                            option.textContent = sheetName;
                            sheetSelect.appendChild(option);
                        });
                        sheetSelect.disabled = false;
                    }
                    if (processBtn) processBtn.disabled = false;
                    showProgress(100, 'فایل آماده پردازش است.');
                } catch (error) {
                    console.error("خطا در بارگذاری فایل اکسل:", error);
                    Swal.fire('خطا', 'فایل اکسل نامعتبر است یا در خواندن آن مشکلی پیش آمده.', 'error');
                    showProgress(0);
                    if (fileNameDisplay) fileNameDisplay.textContent = 'فایل انتخاب نشده...';
                    if (sheetSelect) sheetSelect.innerHTML = '<option value="">- Sheet1 -</option>';
                    if (sheetSelect) sheetSelect.disabled = true;
                    if (processBtn) processBtn.disabled = true;
                }
            };
            reader.readAsArrayBuffer(file);
        });
    }

    if (processBtn) processBtn.addEventListener('click', processData);
    if (exportExcelBtn) exportExcelBtn.addEventListener('click', exportToExcel);
    if (exportPdfBtn) exportPdfBtn.addEventListener('click', exportToPdf);

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

            showProgress(0, 'منتظر انتخاب فایل...');
            Swal.fire('با موفقیت', 'برنامه به حالت اولیه بازگردانده شد.', 'success');
        });
    }
});