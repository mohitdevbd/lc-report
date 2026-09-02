let workbook;
let data = [];
let salesChartInstance = null;
let shareSheetNames = [];
let shareSnapshot = null;
let hiddenCards = [];
let isCustomUpload = false;

const urlParams = new URLSearchParams(window.location.search);
const isShareMode = urlParams.get("share") === "1";
const snapshotPrefix = "snapshot=";
const cardIds = [
    "totalValue",
    "topPerformer",
    "monthlyChart",
    "avgMonthlySales",
    "totalLc",
    "excludedCompanyTotal",
    "excludeNzTopPerformer"
];

const sheetDisplayNameMap = {
    "NZTL": "NZ TEXTILES LTD.",
    "NZAL": "NZ APPARELS LTD.",
    "NZDY": "NZDY FLAX SPINNING LTD."
};

function getSheetDisplayName(sheetName) {
    if (!sheetName) return "";
    const clean = String(sheetName).trim();
    if (sheetDisplayNameMap[clean]) return sheetDisplayNameMap[clean];
    const upper = clean.toUpperCase();
    if (sheetDisplayNameMap[upper]) return sheetDisplayNameMap[upper];
    return clean;
}

const reportState = {
    sheetName:"",
    filterLabel:"All Data",
    rows:[],
    excludedCompanyRows:[],
    excludedTotal:0,
    grandTotal:0,
    avgMonthlySales:0,
    topPerformer:"-",
    topPerformerValue:0,
    totalLC:0,
    excludeNzTopPerformer:"-",
    excludeNzTopPerformerValue:0
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
    sharePanel:document.getElementById("sharePanel"),
    shareTabs:document.getElementById("shareTabs"),
    shareReportTitle:document.getElementById("shareReportTitle"),
    copyShareBtn:document.getElementById("copyShareBtn"),
    grandTotal:document.getElementById("grandTotal"),
    avgMonthlySales:document.getElementById("avgMonthlySales"),
    topPerformer:document.getElementById("topPerformer"),
    topPerformerAmount:document.getElementById("topPerformerAmount"),
    totalLC:document.getElementById("totalLC"),
    excludedCompanyTotal:document.getElementById("excludedCompanyTotal"),
    excludeNzTopPerformer:document.getElementById("excludeNzTopPerformer"),
    excludeNzTopPerformerAmount:document.getElementById("excludeNzTopPerformerAmount"),
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

function getCardElements() {
    return Array.from(document.querySelectorAll("[data-card-id]"));
}

function applyCardVisibility() {
    getCardElements().forEach(card => {
        const cardId = card.dataset.cardId;
        const isHidden = hiddenCards.includes(cardId);
        const button = card.querySelector(".card-visibility-btn");

        card.classList.toggle(
            isShareMode ? "card-hidden" : "is-user-hidden",
            isHidden
        );

        if (isShareMode) {
            card.classList.remove("is-user-hidden");
        }
        else {
            card.classList.remove("card-hidden");
        }

        if (button) {
            button.setAttribute("aria-pressed", String(isHidden));
            button.setAttribute("aria-label", isHidden ? "Show this card" : "Hide from share");
            button.setAttribute("title", isHidden ? "Show this card" : "Hide from share");
            button.innerHTML =
            `<i data-lucide="${isHidden ? "eye-off" : "eye"}"></i>`;
        }
    });

    updateDashboardLayout();

    refreshIcons();

    if (salesChartInstance) {
        window.setTimeout(() => salesChartInstance.resize(), 50);
    }
}

function updateDashboardLayout() {
    const mainGrid = document.querySelector(".dashboard-main-grid");
    const statsColumn = document.querySelector(".stats-left-col");
    const chartCard = document.querySelector('[data-card-id="monthlyChart"]');

    if (!mainGrid || !statsColumn || !chartCard) return;

    const statsCards =
    Array.from(statsColumn.querySelectorAll("[data-card-id]"));
    const visibleStatsCards =
    statsCards.filter(card => !hiddenCards.includes(card.dataset.cardId));
    const chartIsHidden = hiddenCards.includes("monthlyChart");

    statsColumn.classList.toggle(
        "card-hidden",
        isShareMode && visibleStatsCards.length === 0
    );
    mainGrid.classList.toggle(
        "single-column",
        isShareMode && (chartIsHidden || visibleStatsCards.length === 0)
    );
    statsColumn.classList.toggle(
        "stats-grid-mode",
        isShareMode && chartIsHidden && visibleStatsCards.length > 0
    );
}

function toggleCardVisibility(cardId) {
    if (hiddenCards.includes(cardId)) {
        hiddenCards = hiddenCards.filter(item => item !== cardId);
    }
    else {
        hiddenCards = [...hiddenCards, cardId];
    }

    applyCardVisibility();
}

function initCardVisibilityControls() {
    getCardElements().forEach(card => {
        if (card.querySelector(".card-visibility-btn")) return;

        const button = document.createElement("button");

        button.type = "button";
        button.className = "card-visibility-btn";
        button.setAttribute("aria-label", "Hide from share");
        button.addEventListener("click", event => {
            event.stopPropagation();
            toggleCardVisibility(card.dataset.cardId);
        });

        card.appendChild(button);
    });

    applyCardVisibility();
}

function encodeSharePayload(payload) {
    const json = JSON.stringify(payload);
    const binary =
    encodeURIComponent(json)
    .replace(/%([0-9A-F]{2})/g, (match, hex) =>
        String.fromCharCode(parseInt(hex, 16))
    );

    return btoa(binary)
    .replace(/\+/g, "-")
    .replace(/\//g, "_")
    .replace(/=+$/g, "");
}

function decodeSharePayload(encoded) {
    const padded =
    encoded
    .replace(/-/g, "+")
    .replace(/_/g, "/")
    .padEnd(Math.ceil(encoded.length / 4) * 4, "=");
    const json =
    atob(padded)
    .split("")
    .map(character =>
        "%" + character.charCodeAt(0).toString(16).padStart(2, "0")
    )
    .join("");

    return JSON.parse(decodeURIComponent(json));
}

function getSnapshotFromUrl() {
    const hash = window.location.hash.replace(/^#/, "");

    if (!hash.startsWith(snapshotPrefix)) return null;

    try {
        const payload = decodeSharePayload(hash.slice(snapshotPrefix.length));

        return payload && Array.isArray(payload.sheets) ? payload : null;
    }
    catch (error) {
        console.warn("Invalid share snapshot:", error.message);
        return null;
    }
}

function getHiddenCardsFromUrl() {
    const cards = urlParams.get("cards");

    if (!cards) return [];

    return cards
    .split(",")
    .map(card => card.trim())
    .filter(card => cardIds.includes(card));
}

function getSelectedSheetName() {
    return reportState.sheetName || (elements.sheetSelector && elements.sheetSelector.value) || (workbook && workbook.SheetNames && workbook.SheetNames[0]) || "Sheet";
}

function getShareUrl(sheetName = getSelectedSheetName()) {
    const url = new URL(window.location.href);
    const cleanUrl = new URL(url.origin + url.pathname);

    cleanUrl.searchParams.set("share", "1");
    if (sheetName) {
        cleanUrl.searchParams.set("sheet", sheetName);
    }

    if (hiddenCards.length) {
        cleanUrl.searchParams.set("cards", hiddenCards.join(","));
    }

    cleanUrl.hash = "";
    return cleanUrl.toString();
}

function setCopyButtonState(message, icon = "check") {
    if (!elements.copyShareBtn) return;

    elements.copyShareBtn.innerHTML =
    `<i data-lucide="${icon}"></i><span>${message}</span>`;
    refreshIcons();
}

function copyToClipboard(text) {
    if (navigator.clipboard && window.isSecureContext) {
        return navigator.clipboard.writeText(text);
    }

    return new Promise((resolve, reject) => {
        try {
            const textArea = document.createElement("textarea");
            textArea.value = text;
            textArea.style.position = "fixed";
            textArea.style.top = "-9999px";
            textArea.style.left = "-9999px";
            textArea.style.opacity = "0";
            document.body.appendChild(textArea);
            textArea.focus();
            textArea.select();
            const successful = document.execCommand("copy");
            document.body.removeChild(textArea);
            if (successful) {
                resolve();
            } else {
                reject(new Error("Copy command failed"));
            }
        } catch (err) {
            reject(err);
        }
    });
}

function copyShareLink() {
    try {
        const shareUrl = getShareUrl();

        copyToClipboard(shareUrl)
            .then(() => {
                setCopyButtonState("Link Copied!", "check");
                setTimeout(() => {
                    setCopyButtonState("Copy Share Link", "share-2");
                }, 2000);
            })
            .catch(() => {
                window.prompt("Copy this share link:", shareUrl);
            });
    } catch (error) {
        console.error("Error generating share URL:", error);
        window.prompt("Copy this share link:", window.location.href);
    }
}

function syncShareTabs() {
    if (!elements.shareTabs) return;

    const activeSheet = getSelectedSheetName();

    Array.from(elements.shareTabs.querySelectorAll(".share-tab"))
    .forEach(tab => {
        const isActive = tab.dataset.sheet === activeSheet;

        tab.classList.toggle("active", isActive);
        tab.setAttribute("aria-selected", String(isActive));
    });

    if (elements.shareReportTitle) {
        elements.shareReportTitle.textContent = "STATEMENT 2026";
    }
}

function updateShareUrl(sheetName) {
    if (!isShareMode || !sheetName) return;

    const url = new URL(window.location.href);

    url.searchParams.set("share", "1");
    url.searchParams.set("sheet", sheetName);
    window.history.replaceState({}, "", url);
}

function initShareTabs() {
    if (!workbook || !elements.shareTabs || !elements.sharePanel) return;

    shareSheetNames = workbook.SheetNames.slice(0, 3);
    elements.shareTabs.innerHTML = "";

    shareSheetNames.forEach((sheetName, index) => {
        const tab = document.createElement("button");
        const icon = document.createElement("i");
        const label = document.createElement("span");

        tab.type = "button";
        tab.className = "share-tab";
        tab.dataset.sheet = sheetName;
        tab.setAttribute("role", "tab");

        icon.setAttribute("data-lucide", "table-2");
        label.textContent = getSheetDisplayName(sheetName) || ("Sheet " + (index + 1));
        tab.appendChild(icon);
        tab.appendChild(label);

        tab.addEventListener("click", () => {
            if (elements.sheetSelector) {
                elements.sheetSelector.value = sheetName;
            }
            loadSheet(sheetName);
            updateShareUrl(sheetName);
        });

        elements.shareTabs.appendChild(tab);
    });

    elements.sharePanel.classList.toggle("ready", shareSheetNames.length > 0);
    syncShareTabs();
    refreshIcons();
}

function initSnapshotTabs(snapshot) {
    if (!snapshot || !elements.shareTabs || !elements.sharePanel) return false;

    shareSnapshot = snapshot;
    hiddenCards =
    Array.isArray(snapshot.hiddenCards)
    ? snapshot.hiddenCards.filter(card => cardIds.includes(card))
    : getHiddenCardsFromUrl();
    shareSheetNames = snapshot.sheets.map(sheet => sheet.sheetName);
    elements.shareTabs.innerHTML = "";

    snapshot.sheets.forEach((sheet, index) => {
        const tab = document.createElement("button");
        const icon = document.createElement("i");
        const label = document.createElement("span");

        tab.type = "button";
        tab.className = "share-tab";
        tab.dataset.sheet = sheet.sheetName;
        tab.setAttribute("role", "tab");

        icon.setAttribute("data-lucide", "table-2");
        label.textContent = getSheetDisplayName(sheet.sheetName) || ("Sheet " + (index + 1));
        tab.appendChild(icon);
        tab.appendChild(label);

        tab.addEventListener("click", () => {
            renderSnapshotSheet(sheet);
            updateShareUrl(sheet.sheetName);
        });

        elements.shareTabs.appendChild(tab);
    });

    elements.sharePanel.classList.toggle("ready", shareSheetNames.length > 0);

    if (snapshot.generatedAt) {
        elements.todayLabel.textContent = snapshot.generatedAt;
    }

    const requestedSheet = urlParams.get("sheet");
    const initialSheet =
    snapshot.sheets.find(sheet => sheet.sheetName === requestedSheet)
    ||
    snapshot.sheets[0];

    renderSnapshotSheet(initialSheet);
    refreshIcons();

    return true;
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

function getColumnMapForRows(rows) {
    const columns = rows.length ? Object.keys(rows[0]) : [];

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

function calculateMonthlySales(rows, columns) {
    let year = new Date().getFullYear();

    for (const row of rows) {
        const date = parseLCDate(row[columns.lcDate]);

        if (date) {
            year = date.getFullYear();
            break;
        }
    }

    const shortYear = String(year).substring(2);
    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const labels = monthNames.map(month => month + "-" + shortYear);
    const values = Array(12).fill(0);

    rows.forEach(row => {
        const date = parseLCDate(row[columns.lcDate]);

        if (date) {
            values[date.getMonth()] += parseAmount(row[columns.totalValue]);
        }
    });

    return {
        labels,
        values
    };
}

function calculateReportSummary(sheetName, sourceRows) {
    const columns = getColumnMapForRows(sourceRows);

    if (!columns.company || !columns.marketingPerson || !columns.totalValue) {
        return null;
    }

    const reportRows =
    sourceRows.filter(row => hasRequiredReportFields(row, columns));
    const aggregated = aggregateRows(reportRows, columns);
    const monthsInDataset = new Set();

    if (columns.lcDate) {
        reportRows.forEach(row => {
            const date = parseLCDate(row[columns.lcDate]);

            if (date) {
                monthsInDataset.add(
                    date.getFullYear()
                    +
                    "-"
                    +
                    String(date.getMonth() + 1).padStart(2, "0")
                );
            }
        });
    }

    const nzExcludedTotals = {};

    reportRows.forEach(row => {
        const companyName = String(row[columns.company] || "").toUpperCase();

        if (companyName.includes("NZ DENIM") || companyName.includes("NZ FABRIC")) {
            return;
        }

        const marketingPerson = String(row[columns.marketingPerson] || "").trim();

        if (!marketingPerson) return;

        nzExcludedTotals[marketingPerson] =
        (nzExcludedTotals[marketingPerson] || 0)
        +
        parseAmount(row[columns.totalValue]);
    });

    const nzExcludedSorted =
    Object.entries(nzExcludedTotals).sort((a, b) => b[1] - a[1]);
    const monthlySales =
    columns.lcDate
    ? calculateMonthlySales(reportRows, columns)
    : {
        labels:["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
        values:Array(12).fill(0)
    };

    return {
        sheetName,
        filterLabel:"All Data",
        filterDetail:"Uploaded Excel snapshot",
        rows:aggregated.rows,
        excludedCompanyRows:[],
        excludedTotal:0,
        grandTotal:aggregated.grandTotal,
        avgMonthlySales:aggregated.grandTotal / (monthsInDataset.size || 1),
        topPerformer:aggregated.topPerformerName,
        topPerformerValue:aggregated.topPerformerValue,
        totalLC:reportRows.length,
        excludeNzTopPerformer:nzExcludedSorted.length ? nzExcludedSorted[0][0] : "-",
        excludeNzTopPerformerValue:nzExcludedSorted.length ? nzExcludedSorted[0][1] : 0,
        monthlyLabels:monthlySales.labels,
        monthlyValues:monthlySales.values
    };
}

function buildShareSnapshot() {
    if (!workbook) return null;

    const sheets =
    workbook.SheetNames
    .slice(0, 3)
    .map(sheetName => {
        const sheetRows =
        XLSX.utils.sheet_to_json(
            workbook.Sheets[sheetName],
            {
                raw:true,
                defval:""
            }
        );

        return calculateReportSummary(sheetName, sheetRows);
    })
    .filter(Boolean);

    if (!sheets.length) return null;

    return {
        version:1,
        generatedAt:formatDisplayDate(),
        hiddenCards:[...hiddenCards],
        sheets
    };
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

function hasRequiredReportFields(row, columns) {

    const company = String(row[columns.company] || "").trim();
    const marketingPerson = String(row[columns.marketingPerson] || "").trim();

    return (
        company.length > 0 &&
        marketingPerson.length > 0
    );

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
        topPerformerName:rows.length ? rows[0].name : "-",
        topPerformerValue:rows.length ? rows[0].value : 0
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
    elements.avgMonthlySales.textContent = formatMoney(reportState.avgMonthlySales);
    elements.topPerformer.textContent = reportState.topPerformer;
    elements.topPerformerAmount.textContent = reportState.topPerformerValue > 0 ? formatMoney(reportState.topPerformerValue) : "$0.00";
    elements.totalLC.textContent = reportState.totalLC;
    elements.excludedCompanyTotal.textContent = formatMoney(reportState.excludedTotal);
    elements.excludeNzTopPerformer.textContent = reportState.excludeNzTopPerformer;
    elements.excludeNzTopPerformerAmount.textContent = reportState.excludeNzTopPerformerValue > 0 ? formatMoney(reportState.excludeNzTopPerformerValue) : "$0.00";
}

function resetReport() {
    reportState.rows = [];
    reportState.excludedCompanyRows = [];
    reportState.excludedTotal = 0;
    reportState.grandTotal = 0;
    reportState.avgMonthlySales = 0;
    reportState.topPerformer = "-";
    reportState.topPerformerValue = 0;
    reportState.totalLC = 0;
    reportState.excludeNzTopPerformer = "-";
    reportState.excludeNzTopPerformerValue = 0;
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
    const reportRows =
    dateFilteredRows.filter(row => hasRequiredReportFields(row, columns));
    const rowGroups =
    splitIncludedAndExcludedRows(reportRows, columns);
    const aggregated = aggregateRows(rowGroups.includedRows, columns);
    const excludedAggregated =
    aggregateExcludedCompanies(rowGroups.excludedRows, columns);

    // Calculate AVG Monthly Sales
    let activeMonthCount = 1;
    if (filterDetails.type === "month") {
        activeMonthCount = 1;
    } else if (filterDetails.type === "range" && elements.fromDate.value && elements.toDate.value) {
        const start = new Date(elements.fromDate.value);
        const end = new Date(elements.toDate.value);
        const diffMonths = (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth()) + 1;
        activeMonthCount = Math.max(1, diffMonths);
    } else {
        const monthsInDataset = new Set();
        rowGroups.includedRows.forEach(row => {
            const date = parseLCDate(row[columns.lcDate]);
            if (date) {
                const monthStr = date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0");
                monthsInDataset.add(monthStr);
            }
        });
        activeMonthCount = monthsInDataset.size || 1;
    }
    reportState.avgMonthlySales = aggregated.grandTotal / activeMonthCount;

    // Calculate Exclude NZ Denim & NZ Fabric Top Performer
    const nzExcludedTotals = {};
    rowGroups.includedRows.forEach(row => {
        const companyName = String(row[columns.company] || "").toUpperCase();
        if (companyName.includes("NZ DENIM") || companyName.includes("NZ FABRIC")) {
            return; // skip these companies
        }
        const marketingPerson = String(row[columns.marketingPerson] || "").trim();
        if (!marketingPerson) return;
        const value = parseAmount(row[columns.totalValue]);
        nzExcludedTotals[marketingPerson] = (nzExcludedTotals[marketingPerson] || 0) + value;
    });

    const nzExcludedSorted = Object.entries(nzExcludedTotals).sort((a, b) => b[1] - a[1]);
    if (nzExcludedSorted.length > 0) {
        const topName = nzExcludedSorted[0][0];
        const topValue = nzExcludedSorted[0][1];
        reportState.excludeNzTopPerformer = topName;
        reportState.excludeNzTopPerformerValue = topValue;
    } else {
        reportState.excludeNzTopPerformer = "-";
        reportState.excludeNzTopPerformerValue = 0;
    }

    reportState.sheetName = getSelectedSheetName();
    reportState.filterLabel = filterDetails.label;
    reportState.filterDetail = filterDetails.detail;
    reportState.rows = aggregated.rows;
    reportState.excludedCompanyRows = excludedAggregated.rows;
    reportState.excludedTotal = excludedAggregated.total;
    reportState.grandTotal = aggregated.grandTotal;
    reportState.topPerformer = aggregated.topPerformerName;
    reportState.topPerformerValue = aggregated.topPerformerValue;
    reportState.totalLC = rowGroups.includedRows.filter(row => {
        const company = String(row[columns.company] || "").trim();
        const marketing = String(row[columns.marketingPerson] || "").trim();
        return company.length > 0 && marketing.length > 0;
    }).length;

    renderReport();
    updateSalesChart(rowGroups.includedRows, columns);
    applyCardVisibility();

    console.log("Using value column:", columns.totalValue);

    return true;
}

function renderSalesChart(labels, values) {
    const canvas = document.getElementById('monthlySalesChart');
    if (!canvas) return;
    const ctx = canvas.getContext('2d');

    if (salesChartInstance) {
        salesChartInstance.destroy();
    }

    salesChartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Total LC Value ($)',
                data: values,
                backgroundColor: '#797e53',
                hoverBackgroundColor: '#636842',
                borderRadius: 4,
                borderWidth: 0,
                barPercentage: 0.6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    callbacks: {
                        label: function(context) {
                            return 'Value: $' + context.raw.toLocaleString("en-US", {
                                minimumFractionDigits: 2,
                                maximumFractionDigits: 2
                            });
                        }
                    }
                }
            },
            scales: {
                x: {
                    grid: {
                        display: false
                    },
                    ticks: {
                        color: '#64748b',
                        font: {
                            family: 'Plus Jakarta Sans',
                            size: 11,
                            weight: '500'
                        }
                    }
                },
                y: {
                    grid: {
                        color: 'rgba(0, 0, 0, 0.04)'
                    },
                    ticks: {
                        color: '#64748b',
                        font: {
                            family: 'Plus Jakarta Sans',
                            size: 11
                        },
                        callback: function(value) {
                            if (value >= 1e6) {
                                return '$' + (value / 1e6).toFixed(1) + 'M';
                            } else if (value >= 1e3) {
                                return '$' + (value / 1e3).toFixed(0) + 'k';
                            }
                            return '$' + value;
                        }
                    }
                }
            }
        }
    });
}

function updateSalesChart(rows, columns) {
    const monthlySales = calculateMonthlySales(rows, columns);

    renderSalesChart(monthlySales.labels, monthlySales.values);
}

function renderSnapshotSheet(sheet) {
    if (!sheet) return;

    reportState.sheetName = sheet.sheetName;
    reportState.filterLabel = sheet.filterLabel || "All Data";
    reportState.filterDetail = sheet.filterDetail || "Uploaded Excel snapshot";
    reportState.rows = sheet.rows || [];
    reportState.excludedCompanyRows = sheet.excludedCompanyRows || [];
    reportState.excludedTotal = sheet.excludedTotal || 0;
    reportState.grandTotal = sheet.grandTotal || 0;
    reportState.avgMonthlySales = sheet.avgMonthlySales || 0;
    reportState.topPerformer = sheet.topPerformer || "-";
    reportState.topPerformerValue = sheet.topPerformerValue || 0;
    reportState.totalLC = sheet.totalLC || 0;
    reportState.excludeNzTopPerformer = sheet.excludeNzTopPerformer || "-";
    reportState.excludeNzTopPerformerValue = sheet.excludeNzTopPerformerValue || 0;

    renderReport();
    renderSalesChart(
        sheet.monthlyLabels || [],
        sheet.monthlyValues || []
    );
    syncShareTabs();
    applyCardVisibility();
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
        input.addEventListener("change", generateReport);

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
    if (elements.sheetSelector) {
        elements.sheetSelector.value = sheetName;
    }
    loadCompanies();
    generateReport();
    syncShareTabs();
}

function populateSheetSelector() {
    if (!elements.sheetSelector) return;
    elements.sheetSelector.innerHTML = "";

    workbook.SheetNames.forEach(sheet => {
        const option = document.createElement("option");

        option.value = sheet;
        option.textContent = sheet;
        elements.sheetSelector.appendChild(option);
    });
}

function getInitialSheetName() {
    const requestedSheet = urlParams.get("sheet");

    if (requestedSheet && workbook.SheetNames.includes(requestedSheet)) {
        return requestedSheet;
    }

    return workbook.SheetNames[0];
}

function handleExcelUpload(event) {
    const file = event.target.files[0];

    if (!file) return;

    isCustomUpload = true;
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

        populateSheetSelector();
        initShareTabs();
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
if (elements.sheetSelector) {
    elements.sheetSelector.addEventListener("change", event => {
        loadSheet(event.target.value);
        updateShareUrl(event.target.value);
    });
}
if (elements.generateBtn) {
    elements.generateBtn.addEventListener("click", generateReport);
}
elements.exportExcelBtn.addEventListener("click", exportToExcel);
elements.exportPdfBtn.addEventListener("click", exportToPdf);
elements.copyShareBtn.addEventListener("click", copyShareLink);
elements.filterType.addEventListener("change", () => {
    updateFilterVisibility();
    generateReport();
});
elements.monthFilter.addEventListener("change", generateReport);
elements.fromDate.addEventListener("change", generateReport);
elements.toDate.addEventListener("change", generateReport);
elements.toggleCompanies.addEventListener("click", toggleCompanies);

elements.todayLabel.textContent = formatDisplayDate();

if (isShareMode) {
    document.body.classList.add("share-mode");
    document.querySelector(".app-header h1").textContent = "LC Shareable Report";
    hiddenCards = getHiddenCardsFromUrl();
}

updateFilterVisibility();
initCardVisibilityControls();
resetReport();
refreshIcons();
shareSnapshot = getSnapshotFromUrl();

function preloadDefaultExcel() {
    fetch('LC Register 2026.xlsx')
        .then(res => {
            if (!res.ok) throw new Error('Default file not found');
            return res.arrayBuffer();
        })
        .then(ab => {
            workbook = XLSX.read(ab, { type: 'array' });
            populateSheetSelector();
            elements.uploadFileName.textContent = "LC Register 2026.xlsx";
            elements.uploadStatus.textContent = "Auto-loaded";
            elements.uploadCard.classList.add("uploaded");
            initShareTabs();
            loadSheet(getInitialSheetName());
        })
        .catch(err => {
            console.log('Auto-preload status:', err.message);
        });
}

if (!shareSnapshot || !initSnapshotTabs(shareSnapshot)) {
    preloadDefaultExcel();
}
