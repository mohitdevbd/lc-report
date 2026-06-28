let workbook;
let data = [];

const reportState = {
    sheetName:"",
    filterLabel:"All Data",
    rows:[],
    excludedCompanyRows:[],
    excludedTotal:0,
    grandTotal:0,
    topPerformer:"-",
    totalLC:0
};

const elements = {
    excelFile:document.getElementById("excelFile"),
    sheetSelector:document.getElementById("sheetSelector"),
    generateBtn:document.getElementById("generateBtn"),
    exportExcelBtn:document.getElementById("exportExcelBtn"),
    exportPdfBtn:document.getElementById("exportPdfBtn"),
    uploadCard:document.querySelector(".upload-card"),
    uploadFileName:document.getElementById("uploadFileName"),
    uploadStatus:document.getElementById("uploadStatus"),
    todayLabel:document.getElementById("todayLabel"),
    filterType:document.getElementById("filterType"),
    monthBox:document.getElementById("monthBox"),
    monthFilter:document.getElementById("monthFilter"),
    fromBox:document.getElementById("fromBox"),
    fromDate:document.getElementById("fromDate"),
    toBox:document.getElementById("toBox"),
    toDate:document.getElementById("toDate"),
    toggleCompanies:document.getElementById("toggleCompanies"),
    companyWrapper:document.getElementById("companyWrapper"),
    companyList:document.getElementById("companyList"),
    grandTotal:document.getElementById("grandTotal"),
    topPerformer:document.getElementById("topPerformer"),
    totalLC:document.getElementById("totalLC"),
    tableBody:document.querySelector("#reportTable tbody")
};

function normalizeHeader(value) {
    return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function compactHeader(value) {
    return normalizeHeader(value).replace(/\s+/g, "");
}

function findColumn(columns, matchers) {
    for (const matcher of matchers) {

        const column =
        columns.find(item => matcher(normalizeHeader(item), item));

        if (column) return column;

    }

    return undefined;
}

function getColumns() {
    return data.length ? Object.keys(data[0]) : [];
}

function getColumnMap() {
    const columns = getColumns();

    return {
        company:findColumn(
            columns,
            [
                header => header.includes("applicant"),
                header => header.includes("company"),
                header => header.includes("customer")
            ]
        ),
        marketingPerson:findColumn(
            columns,
            [
                header => header.includes("marketing"),
                header => header.includes("sales person"),
                header => header.includes("representative")
            ]
        ),
        totalValue:findColumn(
            columns,
            [
                header => header === "total value",
                header => compactHeader(header) === "totalvalue",
                header => header.includes("lc value"),
                header => header.includes("amount")
            ]
        ),
        lcDate:findColumn(
            columns,
            [
                header => header === "lc date",
                header => header.includes("lc date"),
                header => header.endsWith("date")
            ]
        )
    };
}

function parseLCDate(value) {
    if (!value) return null;

    if (value instanceof Date && !isNaN(value)) {
        return value;
    }

    if (typeof value === "number") {
        const parsed = XLSX.SSF.parse_date_code(value);

        if (!parsed) return null;

        return new Date(
            parsed.y,
            parsed.m - 1,
            parsed.d,
            parsed.H || 0,
            parsed.M || 0,
            parsed.S || 0
        );
    }

    const normalized = String(value).trim();
    const dateParts =
    normalized.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);

    if (dateParts) {
        let year = Number(dateParts[3]);

        if (year < 100) year += 2000;

        const dayFirstDate =
        new Date(
            year,
            Number(dateParts[2]) - 1,
            Number(dateParts[1])
        );

        if (
            dayFirstDate.getFullYear() === year
            &&
            dayFirstDate.getMonth() === Number(dateParts[2]) - 1
            &&
            dayFirstDate.getDate() === Number(dateParts[1])
        ) {
            return dayFirstDate;
        }
    }

    const date = new Date(normalized);

    return isNaN(date) ? null : date;
}

function parseAmount(value) {
    if (typeof value === "number") return value;

    const normalized =
    String(value || "")
    .replace(/,/g, "")
    .trim();

    if (/^\(.+\)$/.test(normalized)) {
        return -parseFloat(
            normalized
            .replace(/[()]/g, "")
            .replace(/[^0-9.-]/g, "")
        ) || 0;
    }

    return parseFloat(
        normalized.replace(/[^0-9.-]/g, "")
    ) || 0;
}

function formatMoney(value) {
    return "$" + Number(value || 0).toLocaleString(
        "en-US",
        {
            minimumFractionDigits:2,
            maximumFractionDigits:2
        }
    );
}

function formatDateForFile(date = new Date()) {
    return [
        date.getFullYear(),
        String(date.getMonth() + 1).padStart(2, "0"),
        String(date.getDate()).padStart(2, "0")
    ].join("-");
}

function formatDisplayDate(date = new Date()) {
    return date.toLocaleDateString(
        "en-US",
        {
            month:"short",
            day:"2-digit",
            year:"numeric"
        }
    );
}

function refreshIcons() {
    if (window.lucide) {
        window.lucide.createIcons();
    }
}

function getSelectedSheetName() {
    return elements.sheetSelector.value || reportState.sheetName || "Sheet";
}

function getFilterDetails() {
    const type = elements.filterType.value;

    if (type === "month") {
        return {
            type,
            label:"Monthly Total",
            detail:elements.monthFilter.value || "Not selected"
        };
    }

    if (type === "range") {
        return {
            type,
            label:"Date Range Total",
            detail:[
                elements.fromDate.value || "Not selected",
                elements.toDate.value || "Not selected"
            ].join(" to ")
        };
    }

    return {
        type,
        label:"All Data",
        detail:"All records"
    };
}

function validateColumns(columns, filterDetails) {
    const missingColumns = [];

    if (!columns.company) missingColumns.push("Applicant/Company");
    if (!columns.marketingPerson) missingColumns.push("Marketing Person");
    if (!columns.totalValue) missingColumns.push("Total Value");
    if (filterDetails.type !== "all" && !columns.lcDate) {
        missingColumns.push("LC Date");
    }

    if (missingColumns.length) {
        alert("Missing required column(s): " + missingColumns.join(", "));
        return false;
    }

    return true;
}

function getExcludedCompanies() {
    return Array.from(
        document.querySelectorAll("#companyList input:checked")
    ).map(input => String(input.value));
}

function getDateFilteredRows(columns, filterDetails) {
    let filteredRows = [...data];

    if (filterDetails.type === "month" && elements.monthFilter.value) {
        filteredRows =
        filteredRows.filter(row => {
            const date = parseLCDate(row[columns.lcDate]);

            if (!date) return false;

            const rowMonth =
            date.getFullYear()
            +
            "-"
            +
            String(date.getMonth() + 1).padStart(2, "0");

            return rowMonth === elements.monthFilter.value;
        });
    }

    if (
        filterDetails.type === "range"
        &&
        elements.fromDate.value
        &&
        elements.toDate.value
    ) {
        const start = new Date(elements.fromDate.value);
        const end = new Date(elements.toDate.value);

        end.setHours(23, 59, 59, 999);

        filteredRows =
        filteredRows.filter(row => {
            const date = parseLCDate(row[columns.lcDate]);

            return date && date >= start && date <= end;
        });
    }

    return filteredRows;
}

function splitIncludedAndExcludedRows(rows, columns) {
    const excludedCompanies = getExcludedCompanies();

    return rows.reduce((result, row) => {
        const company = String(row[columns.company] || "");

        if (excludedCompanies.includes(company)) {
            result.excludedRows.push(row);
        }
        else {
            result.includedRows.push(row);
        }

        return result;
    }, {
        includedRows:[],
        excludedRows:[]
    });
}

function aggregateRows(filteredRows, columns) {
    const totals = {};

    filteredRows.forEach(row => {
        const marketingPerson =
        String(row[columns.marketingPerson] || "").trim();

        if (!marketingPerson) return;

        totals[marketingPerson] =
        (totals[marketingPerson] || 0)
        +
        parseAmount(row[columns.totalValue]);
    });

    const grandTotal =
    Object.values(totals).reduce((sum, value) => sum + value, 0);

    const rows =
    Object.entries(totals)
    .sort((a, b) => b[1] - a[1])
    .map(([name, value]) => ({
        name,
        value,
        contribution:grandTotal > 0 ? (value / grandTotal) * 100 : 0
    }));

    return {
        rows,
        grandTotal,
        topPerformer:rows.length
        ? rows[0].name + " (" + formatMoney(rows[0].value) + ")"
        : "-"
    };
}

function aggregateExcludedCompanies(excludedRows, columns) {
    const totals = {};

    excludedRows.forEach(row => {
        const company =
        String(row[columns.company] || "").trim();

        if (!company) return;

        totals[company] =
        (totals[company] || 0)
        +
        parseAmount(row[columns.totalValue]);
    });

    const rows =
    Object.entries(totals)
    .sort((a, b) => b[1] - a[1])
    .map(([company, value]) => ({
        company,
        value
    }));

    return {
        rows,
        total:rows.reduce((sum, row) => sum + row.value, 0)
    };
}

function appendTableRow(label, value, className) {
    const tr = document.createElement("tr");
    const labelCell = document.createElement("td");
    const valueCell = document.createElement("td");

    if (className) {
        tr.className = className;
    }

    labelCell.textContent = label;
    valueCell.textContent = formatMoney(value);

    tr.appendChild(labelCell);
    tr.appendChild(valueCell);
    elements.tableBody.appendChild(tr);
}

function renderReport() {
    elements.tableBody.innerHTML = "";

    reportState.rows.forEach(row => {
        appendTableRow(row.name, row.value);
    });

    if (reportState.excludedCompanyRows.length) {
        const sectionRow = document.createElement("tr");
        const sectionCell = document.createElement("td");

        sectionRow.className = "excluded-section-row";
        sectionCell.colSpan = 2;
        sectionCell.textContent = "Excluded Companies Summary";
        sectionRow.appendChild(sectionCell);
        elements.tableBody.appendChild(sectionRow);

        reportState.excludedCompanyRows.forEach(row => {
            appendTableRow(
                row.company,
                row.value,
                "excluded-company-row"
            );
        });

        appendTableRow(
            "Excluded Companies Total",
            reportState.excludedTotal,
            "excluded-total-row"
        );
    }

    elements.grandTotal.textContent = formatMoney(reportState.grandTotal);
    elements.topPerformer.textContent = reportState.topPerformer;
    elements.totalLC.textContent = reportState.totalLC;
}

function resetReport() {
    reportState.rows = [];
    reportState.excludedCompanyRows = [];
    reportState.excludedTotal = 0;
    reportState.grandTotal = 0;
    reportState.topPerformer = "-";
    reportState.totalLC = 0;
    reportState.filterLabel = "All Data";
    renderReport();
}

function generateReport() {
    if (!data.length) {
        alert("Please Upload Excel File");
        return false;
    }

    const columns = getColumnMap();
    const filterDetails = getFilterDetails();

    if (!validateColumns(columns, filterDetails)) {
        return false;
    }

    const dateFilteredRows = getDateFilteredRows(columns, filterDetails);
    const rowGroups =
    splitIncludedAndExcludedRows(dateFilteredRows, columns);
    const aggregated = aggregateRows(rowGroups.includedRows, columns);
    const excludedAggregated =
    aggregateExcludedCompanies(rowGroups.excludedRows, columns);

    reportState.sheetName = getSelectedSheetName();
    reportState.filterLabel = filterDetails.label;
    reportState.filterDetail = filterDetails.detail;
    reportState.rows = aggregated.rows;
    reportState.excludedCompanyRows = excludedAggregated.rows;
    reportState.excludedTotal = excludedAggregated.total;
    reportState.grandTotal = aggregated.grandTotal;
    reportState.topPerformer = aggregated.topPerformer;
    reportState.totalLC = rowGroups.includedRows.length;

    renderReport();

    console.log("Using value column:", columns.totalValue);

    return true;
}

function loadCompanies() {
    elements.companyList.innerHTML = "";

    if (!data.length) return;

    const columns = getColumnMap();

    if (!columns.company) return;

    const companies =
    [...new Set(data.map(row => row[columns.company]).filter(Boolean))]
    .sort();

    companies.forEach(company => {
        const item = document.createElement("div");
        const label = document.createElement("label");
        const input = document.createElement("input");

        input.type = "checkbox";
        input.value = company;

        label.appendChild(input);
        label.appendChild(document.createTextNode(" " + company));
        item.appendChild(label);
        elements.companyList.appendChild(item);
    });
}

function loadSheet(sheetName) {
    if (!workbook || !sheetName) return;

    const sheet = workbook.Sheets[sheetName];

    data =
    XLSX.utils.sheet_to_json(
        sheet,
        {
            raw:true,
            defval:""
        }
    );

    reportState.sheetName = sheetName;
    loadCompanies();
    resetReport();
}

function handleExcelUpload(event) {
    const file = event.target.files[0];

    if (!file) return;

    elements.uploadFileName.textContent = file.name;
    elements.uploadStatus.textContent = "File Uploaded";
    elements.uploadCard.classList.add("uploaded");

    const reader = new FileReader();

    reader.onload = function(evt) {
        workbook =
        XLSX.read(
            evt.target.result,
            {
                type:"array"
            }
        );

        elements.sheetSelector.innerHTML = "";

        workbook.SheetNames.forEach(sheet => {
            const option = document.createElement("option");

            option.value = sheet;
            option.textContent = sheet;
            elements.sheetSelector.appendChild(option);
        });

        loadSheet(workbook.SheetNames[0]);
    };

    reader.readAsArrayBuffer(file);
}

function updateFilterVisibility() {
    elements.monthBox.style.display = "none";
    elements.fromBox.style.display = "none";
    elements.toBox.style.display = "none";

    if (elements.filterType.value === "month") {
        elements.monthBox.style.display = "flex";
    }

    if (elements.filterType.value === "range") {
        elements.fromBox.style.display = "flex";
        elements.toBox.style.display = "flex";
    }
}

function ensureReportForExport() {
    return generateReport();
}

function getExportFileName(extension) {
    return [
        "LC_Report",
        getSelectedSheetName(),
        formatDateForFile()
    ].join("_") + "." + extension;
}

function exportToExcel() {
    if (!ensureReportForExport()) return;

    const exportRows =
    reportState.rows.map(row => ({
        "Marketing Person":row.name,
        "Total LC Value":Number(row.value.toFixed(2))
    }));

    exportRows.push({
        "Marketing Person":"Grand Total",
        "Total LC Value":Number(reportState.grandTotal.toFixed(2))
    });

    if (reportState.excludedCompanyRows.length) {
        exportRows.push({
            "Marketing Person":"",
            "Total LC Value":""
        });

        exportRows.push({
            "Marketing Person":"Excluded Companies Summary",
            "Total LC Value":""
        });

        reportState.excludedCompanyRows.forEach(row => {
            exportRows.push({
                "Marketing Person":row.company,
                "Total LC Value":Number(row.value.toFixed(2))
            });
        });

        exportRows.push({
            "Marketing Person":"Excluded Companies Total",
            "Total LC Value":Number(reportState.excludedTotal.toFixed(2))
        });
    }

    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const exportWorkbook = XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(exportWorkbook, worksheet, "LC Report");
    XLSX.writeFile(exportWorkbook, getExportFileName("xlsx"));
}

function exportToPdf() {
    if (!ensureReportForExport()) return;

    const jsPDF = window.jspdf && window.jspdf.jsPDF;

    if (!jsPDF) {
        alert("PDF library is not loaded. Please check your internet connection and reload.");
        return;
    }

    const doc = new jsPDF();
    const generatedDate = formatDateForFile();

    doc.setFontSize(16);
    doc.text("LC Statement Report", 14, 16);

    doc.setFontSize(10);
    doc.text("Selected Sheet: " + getSelectedSheetName(), 14, 25);
    doc.text("Applied Filter: " + reportState.filterLabel, 14, 31);
    doc.text("Filter Detail: " + reportState.filterDetail, 14, 37);
    doc.text("Generated Date: " + generatedDate, 14, 43);

    const pdfBody =
    reportState.rows.map(row => [
        row.name,
        formatMoney(row.value)
    ]);

    pdfBody.push([
        "Grand Total",
        formatMoney(reportState.grandTotal)
    ]);

    if (reportState.excludedCompanyRows.length) {
        pdfBody.push([
            "Excluded Companies Summary",
            ""
        ]);

        reportState.excludedCompanyRows.forEach(row => {
            pdfBody.push([
                row.company,
                formatMoney(row.value)
            ]);
        });

        pdfBody.push([
            "Excluded Companies Total",
            formatMoney(reportState.excludedTotal)
        ]);
    }

    doc.autoTable({
        startY:50,
        head:[["Marketing Person", "Total LC Value"]],
        body:pdfBody,
        theme:"grid",
        headStyles:{
            fillColor:[37,99,235]
        }
    });

    const finalY = doc.lastAutoTable.finalY + 10;

    doc.setFontSize(11);
    doc.text("Grand Total: " + formatMoney(reportState.grandTotal), 14, finalY);

    doc.save(getExportFileName("pdf"));
}

function toggleCompanies() {
    const isOpen = elements.companyWrapper.classList.toggle("open");

    elements.toggleCompanies.innerHTML =
    '<i data-lucide="list-filter"></i>' +
    (isOpen ? "Hide Exclude Companies" : "Show Exclude Companies");

    elements.toggleCompanies.setAttribute("aria-expanded", String(isOpen));
    refreshIcons();
}

elements.excelFile.addEventListener("change", handleExcelUpload);
elements.sheetSelector.addEventListener("change", event => loadSheet(event.target.value));
elements.generateBtn.addEventListener("click", generateReport);
elements.exportExcelBtn.addEventListener("click", exportToExcel);
elements.exportPdfBtn.addEventListener("click", exportToPdf);
elements.filterType.addEventListener("change", updateFilterVisibility);
elements.toggleCompanies.addEventListener("click", toggleCompanies);

elements.todayLabel.textContent = formatDisplayDate();
updateFilterVisibility();
resetReport();
refreshIcons();
