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
        chartsContainer.innerHTML = '<p id="noChartsMessage" style="text-align: center; color: #777; margin-top: 20px;">برای نمایش نمودارها، لطفاً ابتدا داده‌ها را پردازش کنید.</p>'; // پاک کردن محتوای HTML کانتینر نمودارها و بازگرداندن پیام
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

    // تعریف اندازه ثابت برای همه نمودارها
    const chartWidth = '100%';
    const chartHeight = '250px';

    parsedData.forEach(customer => {
        const chartId = `chart-${customer.bodyNumber}`; // ID منحصر به فرد برای هر نمودار
        const chartDiv = document.createElement('div');
        chartDiv.className = 'chart-container my-3 p-3 border rounded shadow-sm bg-transparent text-white';
        chartDiv.id = `chart-div-${customer.bodyNumber}`; // ID برای اسکرول کردن به آن
        chartDiv.innerHTML = `
            <h3 class="text-primary">نمودار بار مصرفی مشترک: ${customer.customerName} (شماره بدنه: ${customer.bodyNumber})</h3>
            <p class="text-white-50 address-font-small">آدرس: ${customer.address}</p>
            <p class="text-info" style="margin-top: 0; margin-bottom: 5px;">کاهش بار: ${customer.reductionKW} KW &nbsp; | &nbsp; درصد کاهش: ${customer.reductionPercent}%</p>
            <div style="width: ${chartWidth}; height: ${chartHeight}; margin: auto;">
                <canvas id="${chartId}"></canvas>
            </div>
            <hr class="text-white-50">
        `;
        if (chartsContainer) chartsContainer.appendChild(chartDiv);

        const ctx = document.getElementById(chartId);
        if (ctx) {
            const newChart = new Chart(ctx.getContext('2d'), {
                type: 'line',
                data: {
                    labels: customer.timeLabels, // These are now H1, H2, ... H24
                    datasets: [{
                        label: 'بار مصرفی (KW)',
                        data: customer.loadProfileData,
                        borderColor: 'rgb(75, 192, 192)', // Keep this for line, or use gradient for fill
                        backgroundColor: 'rgba(75, 192, 192, 0.5)', // Added for area chart effect
                        tension: 0.4, // Smoother lines
                        fill: true, // Fill the area under the line
                        pointRadius: 0 // نقاط روی نمودار نمایش داده نشوند
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    scales: {
                        x: {
                            title: {
                                display: true,
                                text: 'زمان',
                                color: 'black'
                            },
                            ticks: {
                                color: 'black'
                            },
                            grid: {
                                color: 'rgba(0, 0, 0, 0.2)'
                            }
                        },
                        y: {
                            title: {
                                display: true,
                                text: 'میزان مصرف مشترک (KW)',
                                color: 'black'
                            },
                            ticks: {
                                color: 'black'
                            },
                            grid: {
                                color: 'rgba(0, 0, 0, 0.2)'
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
                            display: false, // Removed legend
                        },
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
                                            size: 10,
                                            color: 'white'
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
    // حذف از آرایه parsedData
    const initialLength = parsedData.length;
    parsedData = parsedData.filter(customer => customer.id !== customerId);

    if (parsedData.length === initialLength) {
        console.warn(`مشتری با ID ${customerId} در parsedData یافت نشد.`);
        // نیازی به Swal.fire نیست چون قرار است بی صدا حذف کند.
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
    // از customerId که همان `customer-${bodyNumber}` است، bodyNumber را استخراج می‌کنیم
    const bodyNumber = customerId.split('-')[1];
    const chartDivToRemove = document.getElementById(`chart-div-${bodyNumber}`);

    if (chartDivToRemove) {
        const canvasId = `chart-${bodyNumber}`;
        const chartInstanceIndex = currentCharts.findIndex(chart => chart.canvas.id === canvasId);

        if (chartInstanceIndex !== -1) {
            currentCharts[chartInstanceIndex].destroy(); // تخریب نمونه نمودار
            currentCharts.splice(chartInstanceIndex, 1); // حذف از آرایه
            console.log(`نمودار برای مشتری ${bodyNumber} حذف شد.`);
        }
        chartDivToRemove.remove(); // حذف عنصر div نمودار از DOM
    } else {
        console.warn(`نمودار (div) برای مشتری ${bodyNumber} در DOM یافت نشد.`);
    }

    // به‌روزرسانی وضعیت UI در صورت خالی شدن داده‌ها
    if (parsedData.length === 0) {
        const resultsTableBody = document.querySelector('#resultsTable tbody');
        if (resultsTableBody) {
            resultsTableBody.innerHTML = '';
            const row = resultsTableBody.insertRow();
            const cell = row.insertCell();
            cell.colSpan = 13;
            cell.textContent = 'هیچ داده‌ای بر اساس فیلترهای اعمال شده یافت نشد.';
            cell.style.textAlign = 'center';
        }
        const noChartsMessage = document.getElementById('noChartsMessage');
        if (noChartsMessage) noChartsMessage.style.display = 'block';
        const exportExcelBtn = document.getElementById('exportExcelBtn');
        const exportPdfBtn = document.getElementById('exportPdfBtn');
        const exportChartsAsImagesBtn = document.getElementById('exportChartsAsImagesBtn');
        if (exportExcelBtn) exportExcelBtn.disabled = true;
        if (exportPdfBtn) exportPdfBtn.disabled = true;
        if (exportChartsAsImagesBtn) exportChartsAsImagesBtn.disabled = true;
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

    const wsData = [];

    // اضافه کردن عنوان و هدر (به صورت متنی، افزودن عکس مستقیماً از اینجا امکان‌پذیر نیست)
    // ساخت یک عنوان که بتواند در یک ردیف ادغام شود.
    wsData.push(["", "", "گزارش تحلیل بار مشترکین", "", "", "", "", "", "", "", ""]); // عنوان اصلی
    wsData.push([""]); // ردیف خالی برای فاصله

    const headerRow = [
        "ردیف", "شماره بدنه", "نام مشترک", "شناسه قبض", "آدرس مشترک",
        "شماره اشتراک", "دیماند قراردادی (KW)", "بار صبح (KW)",
        "بار عصر (KW)", "میزان کاهش (KW)", "درصد کاهش (%)"
    ];
    wsData.push(headerRow);

    parsedData.forEach(customer => {
        wsData.push([
            customer.rowNum,
            customer.bodyNumber,
            customer.customerName,
            customer.billId,
            customer.address,
            customer.subscriptionNumber,
            parseFloat(customer.contractDemand), // اطمینان از عدد
            parseFloat(customer.morningLoad),
            parseFloat(customer.eveningLoad),
            parseFloat(customer.reductionKW),
            parseFloat(customer.reductionPercent)
        ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // ادغام سلول برای عنوان اصلی
    if (wsData.length > 0) {
        ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: headerRow.length - 1 } }];
        // تنظیم استایل برای عنوان (مرکزچین و بزرگتر)
        if (ws['A1']) {
            ws['A1'].s = {
                font: { sz: 16, bold: true, name: "Vazirmatn" },
                alignment: { horizontal: "center", vertical: "center" }
            };
        }
    }


    // تنظیم عرض ستون‌ها (تلاش برای فیت کردن)
    const colWidths = headerRow.map((_, i) => {
        let maxLen = headerRow[i].length;
        // از ردیف بعد از هدر شروع کنید
        for (let j = 2; j < wsData.length; j++) { // Adjusted starting row to consider header at row 2
            const cellValue = wsData[j][i];
            if (cellValue !== undefined && cellValue !== null) {
                // برای اعداد، طول نمایش را در نظر بگیرید
                const stringValue = typeof cellValue === 'number' ? String(cellValue) : cellValue;
                maxLen = Math.max(maxLen, String(stringValue).length);
            }
        }
        return { wch: maxLen + 2 }; // +2 for padding
    });
    ws['!cols'] = colWidths;

    // اعمال برخی استایل‌های اولیه برای هدر جدول و داده‌ها
    // این قسمت به دلیل محدودیت‌های XLSX.utils به صورت پیشرفته قابل پیاده‌سازی نیست
    // و نیازمند کتابخانه‌های پیچیده‌تر استایل‌دهی اکسل است (مثلاً xlsx-js-style).
    // برای مثال:
    // for (let i = 0; i < headerRow.length; i++) {
    //     const cellRef = XLSX.utils.encode_cell({ r: 2, c: i }); // r=2 for header row
    //     if (ws[cellRef]) {
    //         ws[cellRef].s = {
    //             fill: { fgColor: { rgb: "FFE0E0E0" } }, // Light gray background
    //             font: { bold: true },
    //             alignment: { horizontal: "center" },
    //             border: {
    //                 top: { style: "thin", color: { auto: 1 } },
    //                 bottom: { style: "thin", color: { auto: 1 } },
    //                 left: { style: "thin", color: { auto: 1 } },
    //                 right: { style: "thin", color: { auto: 1 } }
    //             }
    //         };
    //     }
    // }

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

    // **مهم:** برای استفاده از فونت فارسی در jsPDF، باید فونت را به صورت Base64 اضافه کنید.
    // دستور شما: doc.addFont('./fonts/Amiri-Regular.ttf', 'Amiri', 'normal');
    // برای افزودن به صورت محلی و با `addFont` باید فایل فونت را به Base64 تبدیل کرده و اینجا قرار دهید.
    // یا مطمئن شوید که تابع `addAmiriFont` در فایل جداگانه (مثلاً `amiri-font.js`) قبل از این اسکریپت لود شده است.
    // با دستور `doc.addFont('./fonts/Amiri-Regular.ttf', 'Amiri', 'normal');` اگر در یک محیط سرور Node.js یا electron باشید و فونت در فایل سیستم در دسترس باشد کار می‌کند.
    // در محیط مرورگر، نیاز به فونت به صورت Base64 یا بارگذاری از یک URL قابل دسترس (و معمولاً با CORS تنظیم شده) دارید.
    try {
        // فرض می‌کنیم که این فونت به درستی از طریق `addFont` شما در `Script_Charts10.js` اضافه شده است.
        // اگر این خط باعث خطا می‌شود، ممکن است نیاز به تبدیل فونت به Base64 و افزودن آن به کد داشته باشید.
        doc.addFont('./fonts/Amiri-Regular.ttf', 'Amiri', 'normal');
        doc.setFont('Amiri');
        doc.setR2L(true); // برای راست به چپ کردن متن
    } catch (e) {
        console.warn("خطا در افزودن فونت فارسی Amiri. ممکن است متن فارسی به درستی نمایش داده نشود.", e);
        Swal.fire('هشدار', 'فونت فارسی برای PDF یافت نشد. ممکن است متن به درستی نمایش داده نشود.', 'warning');
        doc.setFont('helvetica'); // Fallback font
        doc.setR2L(false);
    }

    // ===========================================================================
    // اضافه کردن عکس header.png به بالای هر صفحه PDF
    // ===========================================================================
    // برای این کار، شما نیاز دارید که header.png به صورت Base64 در دسترس باشد
    // یا از یک URL قابل دسترس بارگذاری شود (که نیاز به سرور و CORS دارد).
    // ساده‌ترین راه برای مرورگرها، تبدیل عکس به Base64 است.
    // من یک placeholder برای این کار قرار می‌دهم. شما باید 'BASE64_IMAGE_STRING_HERE' را با مقدار واقعی جایگزین کنید.
    // می توانید از ابزارهای آنلاین برای تبدیل header.png به Base64 استفاده کنید.
    const headerImage = new Image();
    headerImage.src = 'header.png'; // مسیر عکس
    await new Promise(resolve => {
        headerImage.onload = resolve;
        headerImage.onerror = () => {
            console.warn("خطا در بارگذاری header.png برای PDF. تصویر اضافه نخواهد شد.");
            resolve(); // ادامه حتی در صورت خطا
        };
    });

    const addHeaderToPdf = (docInstance) => {
        if (headerImage.complete && headerImage.naturalWidth > 0) {
            const imgWidth = 550; // عرض تقریبی عکس در PDF
            const imgHeight = (headerImage.naturalHeight * imgWidth) / headerImage.naturalWidth;
            const xPos = (docInstance.internal.pageSize.getWidth() - imgWidth) / 2;
            docInstance.addImage(headerImage, 'PNG', xPos, 10, imgWidth, imgHeight); // 10pt از بالا
        } else {
            // اگر عکس بارگذاری نشد، یک عنوان متنی ساده قرار دهید
            docInstance.setFontSize(14);
            docInstance.setFont('Amiri');
            docInstance.setR2L(true);
            docInstance.text("گزارش تحلیل بار مشترکین", docInstance.internal.pageSize.getWidth() / 2, 30, { align: "center" });
        }
    };
    // ===========================================================================


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
        startY: 100, // فضای کافی برای هدر عکس در نظر گرفته شد
        theme: 'grid',
        styles: { font: 'Amiri', fontStyle: 'normal', halign: 'center', cellPadding: 5, fontSize: 8 },
        headStyles: { fillColor: [22, 160, 133], fontSize: 9 },
        margin: { top: 70, right: 30, left: 30 }, // حاشیه بالا برای هدر
        didDrawPage: function (data) {
            addHeaderToPdf(doc); // افزودن هدر به هر صفحه
            // اضافه کردن شماره صفحه
            doc.setFontSize(8);
            doc.setFont('helvetica'); // فونت انگلیسی برای شماره صفحه
            doc.setR2L(false); // چپ به راست برای شماره صفحه
            const pageNumber = doc.internal.getNumberOfPages();
            doc.text(`Page ${pageNumber}`, doc.internal.pageSize.getWidth() - 40, doc.internal.pageSize.getHeight() - 20, { align: 'right' });
        }
    });

    let yOffset = doc.autoTable.previous.finalY + 30;

    Swal.fire({
        title: 'در حال تولید PDF...',
        html: 'لطفاً منتظر بمانید تا نمودارها به PDF اضافه شوند.',
        allowOutsideClick: false,
        didOpen: () => {
            Swal.showLoading();
        }
    });

    for (const customer of parsedData) {
        const chartId = `chart-${customer.bodyNumber}`;
        const canvas = document.getElementById(chartId);

        if (canvas) {
            try {
                // اطمینان از رندر کامل نمودار
                await new Promise(resolve => setTimeout(resolve, 500)); // افزایش زمان برای اطمینان بیشتر

                // افزایش مقیاس رندر Canvas برای کیفیت بهتر در PDF
                const scale = 2; // مثلاً دو برابر کیفیت
                const tempCanvas = document.createElement('canvas');
                tempCanvas.width = canvas.width * scale;
                tempCanvas.height = canvas.height * scale;
                const tempCtx = tempCanvas.getContext('2d');
                tempCtx.fillStyle = 'white'; // پس‌زمینه سفید برای نمودار
                tempCtx.fillRect(0, 0, tempCanvas.width, tempCanvas.height);
                tempCtx.drawImage(canvas, 0, 0, tempCanvas.width, tempCanvas.height);


                const imgData = tempCanvas.toDataURL('image/png', 1.0);
                const imgWidth = 500; // عرض تصویر در PDF
                const imgHeight = (tempCanvas.height * imgWidth) / tempCanvas.width; // حفظ نسبت ابعاد
                const margin = 40;
                const availableWidth = doc.internal.pageSize.getWidth() - 2 * margin;

                // بررسی می‌کنیم آیا نمودار جدید در صفحه فعلی جا می‌شود یا نه
                if (yOffset + imgHeight + 80 > doc.internal.pageSize.getHeight() - margin) {
                    doc.addPage();
                    yOffset = margin + 50; // شروع از پایین هدر
                    addHeaderToPdf(doc); // افزودن هدر به صفحه جدید
                    doc.setFontSize(8); // تنظیم مجدد فونت برای شماره صفحه
                    doc.setFont('helvetica');
                    doc.setR2L(false);
                    const pageNumber = doc.internal.getNumberOfPages();
                    doc.text(`Page ${pageNumber}`, doc.internal.pageSize.getWidth() - 40, doc.internal.pageSize.getHeight() - 20, { align: 'right' });
                }

                doc.setFont('Amiri'); // اطمینان از فونت فارسی
                doc.setR2L(true);
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
                // Swal.fire('خطا', `مشکل در افزودن نمودار مشتری ${customer.customerName} به PDF.`, 'error');
                // نیازی به نمایش ارور مجدد نیست، چون Swal در حال نمایش است و در نهایت باید بسته شود.
            }
        } else {
            console.warn(`عنصر Canvas برای نمودار مشتری ${customer.bodyNumber} در DOM یافت نشد.`);
        }
    }

    doc.save("گزارش_تحلیل_بار.pdf");
    Swal.close(); // بستن SweetAlert پس از اتمام
    console.log("خروجی PDF ایجاد شد.");
}

/**
 * تابع برای خروجی گرفتن تمامی نمودارها به صورت فایل تصویری PNG
 * و فشرده‌سازی آن‌ها در یک فایل ZIP.
 */
async function exportChartsAsImages() {
    if (parsedData.length === 0) {
        Swal.fire('هشدار', 'داده‌ای برای خروجی عکس نمودارها وجود ندارد.', 'warning');
        return;
    }

    // استفاده از window.JSZip به جای import
    const JSZip = window.JSZip;
    if (!JSZip) {
        Swal.fire('خطا', 'کتابخانه JSZip یافت نشد. لطفاً مطمئن شوید که اسکریپت آن به درستی بارگذاری شده است.', 'error');
        console.error("JSZip is not defined. Ensure it's loaded before this function is called.");
        return;
    }

    const zip = new JSZip();
    const date = new Date();
    const folderName = `نمودارهای_مصرف_${date.toLocaleDateString('fa-IR').replace(/\//g, '-')}_${date.getHours().toString().padStart(2, '0')}-${date.getMinutes().toString().padStart(2, '0')}-${date.getSeconds().toString().padStart(2, '0')}`;
    const imagesFolder = zip.folder(folderName);

    Swal.fire({
        title: 'در حال تولید تصاویر نمودارها...',
        html: 'لطفاً منتظر بمانید.',
        allowOutsideClick: false,
        didOpen: () => {
            Swal.showLoading();
        }
    });

    for (const customer of parsedData) {
        const chartDiv = document.getElementById(`chart-div-${customer.bodyNumber}`);
        if (chartDiv) {
            try {
                // استفاده از html2canvas برای گرفتن عکس از div شامل عنوان و نمودار
                const canvas = await html2canvas(chartDiv, {
                    scale: 2, // افزایش مقیاس برای کیفیت بهتر
                    backgroundColor: 'rgba(255, 255, 255, 0.8)', // استفاده از پس‌زمینه کارت
                    useCORS: true,
                    allowTaint: true,
                    logging: false // کاهش لاگ‌ها
                });

                const imgData = canvas.toDataURL('image/png');
                const imgBase64 = imgData.split(',')[1]; // فقط داده Base64

                const fileName = `مشتری_${customer.bodyNumber}_${customer.customerName}.png`;
                imagesFolder.file(fileName, imgBase64, { base64: true });
                console.log(`عکس نمودار مشتری ${customer.bodyNumber} به ZIP اضافه شد.`);

            } catch (error) {
                console.error(`خطا در گرفتن عکس از نمودار مشتری ${customer.bodyNumber}:`, error);
                // نیازی به نمایش ارور مجدد نیست، چون Swal در حال نمایش است و در نهایت باید بسته شود.
            }
        } else {
            console.warn(`عنصر نمودار (div) برای مشتری ${customer.bodyNumber} یافت نشد.`);
        }
    }

    try {
        const content = await zip.generateAsync({ type: "blob" });
        const date = new Date();
        const fileName = `نمودارهای_مصرف_${date.getFullYear()}-${(date.getMonth() + 1).toString().padStart(2, '0')}-${date.getDate().toString().padStart(2, '0')}_${date.getHours().toString().padStart(2, '0')}-${date.getMinutes().toString().padStart(2, '0')}-${date.getSeconds().toString().padStart(2, '0')}.zip`;

        const a = document.createElement('a');
        document.body.appendChild(a);
        a.style.display = 'none';
        const url = window.URL.createObjectURL(content);
        a.href = url;
        a.download = fileName;
        a.click();
        window.URL.revokeObjectURL(url);
        Swal.fire('موفقیت', 'فایل ZIP نمودارها با موفقیت دانلود شد.', 'success');
        console.log("فایل ZIP نمودارها ایجاد و دانلود شد.");

    } catch (zipError) {
        console.error("خطا در فشرده‌سازی یا دانلود فایل ZIP:", zipError);
        Swal.fire('خطا', 'مشکل در تولید یا دانلود فایل ZIP نمودارها.', 'error');
    } finally {
        Swal.close();
    }
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
    const exportChartsAsImagesBtn = document.getElementById('exportChartsAsImagesBtn');
    const resetAppBtn = document.getElementById('resetAppBtn');
    const fileNameDisplay = document.getElementById('fileNameDisplay');

    const chkEvening = document.getElementById('chkEvening');
    const txtEvening = document.getElementById('txtEvening');
    const chkReduction = document.getElementById('chkReduction');
    const txtReduction = document.getElementById('txtReduction');

    showProgress(0);

    if (sheetSelect) sheetSelect.disabled = true;
    if (processBtn) processBtn.disabled = true;
    if (exportExcelBtn) exportExcelBtn.disabled = true;
    if (exportPdfBtn) exportPdfBtn.disabled = true;
    if (exportChartsAsImagesBtn) exportChartsAsImagesBtn.disabled = true; // غیرفعال کردن دکمه جدید
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
    if (exportChartsAsImagesBtn) exportChartsAsImagesBtn.addEventListener('click', exportChartsAsImages); // افزودن event listener

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
            if (exportChartsAsImagesBtn) exportChartsAsImagesBtn.disabled = true; // غیرفعال کردن دکمه جدید
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