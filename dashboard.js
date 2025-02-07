let weekData = []; // Holds Excel data
const categories = ["Strip", "Pot", "Herbs"]; // Define all categories dynamically

// Load Excel file
function loadExcel() {
    const input = document.getElementById('excelFileInput');
    const file = input?.files[0];

    if (file) {
        const reader = new FileReader();

        reader.onload = function (e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });

            // Assume data is in the first sheet
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];

            // Convert sheet to JSON
            weekData = XLSX.utils.sheet_to_json(worksheet);

            if (weekData.length > 0) {
                populateWeekSelector();
                updateDashboard();
            } else {
                console.warn('No data available in the Excel sheet.');
            }
        };

        reader.readAsArrayBuffer(file);
    } else {
        console.error('No file selected or invalid file input.');
    }
}

// Populate week selector dropdown
function populateWeekSelector() {
    const weekSelector = document.getElementById('week');
    if (!weekSelector) {
        console.error('Week selector element not found.');
        return;
    }

    weekSelector.innerHTML = ''; // Clear previous options

    // Add options for available weeks
    weekData.forEach(row => {
        if (row.Week) {
            const option = document.createElement('option');
            option.value = row.Week;
            option.textContent = `Week ${row.Week}`;
            weekSelector.appendChild(option);
        }
    });
}

// Update dashboard for all sections
function updateDashboard() {
    const weekSelector = document.getElementById('week');
    const selectedWeek = parseInt(weekSelector?.value);

    if (selectedWeek && weekData.length > 0) {
        categories.forEach(category => updateCategory(category, selectedWeek));
        updateSection2(selectedWeek);
        updateSection3(selectedWeek);
    } else {
        console.warn('Cannot update dashboard: No week selected or data unavailable.');
    }
}

// Update cards for a specific category
function updateCategory(category, selectedWeek) {
    const previousWeek = selectedWeek - 1;
    let budgetToDate = 0, actualToDate = 0;

    // Sum values up to the previous week (excluding the selected week)
    weekData.forEach(row => {
        const week = parseInt(row.Week);
        if (week <= previousWeek) {
            budgetToDate += parseInt(row[`${category}_Budget`] || 0);
            actualToDate += parseInt(row[`${category}_Actual`] || 0);
        }
    });

    // Calculate variance
    const variance = actualToDate - budgetToDate;

    // Update DOM elements dynamically
    updateElement(`${category.toLowerCase()}-budget`, budgetToDate);
    updateElement(`${category.toLowerCase()}-actual`, actualToDate);
    updateElement(`${category.toLowerCase()}-variance`, variance);
}


// Update Section 2 with previous week figures
function updateSection2(selectedWeek) {
    const previousWeek = selectedWeek - 1;

    if (previousWeek < 1) {
        clearPreviousWeekData();
        return;
    }

    const rowData = weekData.find(row => parseInt(row.Week) === previousWeek);

    if (!rowData) {
        clearPreviousWeekData();
        return;
    }

    categories.forEach(category => {
        const budget = parseInt(rowData[`${category}_Budget`] || 0);
        const actual = parseInt(rowData[`${category}_Actual`] || 0);
        const ordersLastYear = parseInt(rowData[`${category}_Orders_Last_Year`] || 0);
        const ordersThisYear = parseInt(rowData[`${category}_Orders_This_Year`] || 0);

        updateElement(`prev-${category.toLowerCase()}-budget`, budget);
        updateElement(`prev-${category.toLowerCase()}-actual`, actual);
        updateElement(`prev-${category.toLowerCase()}-orders-last-year`, ordersLastYear);
        updateElement(`prev-${category.toLowerCase()}-orders-this-year`, ordersThisYear);
    });
}


// Update Section 3 with specific data for the current week
function updateSection3(selectedWeek) {
    const rowData = weekData.find(row => parseInt(row.Week) === selectedWeek);

    if (!rowData) {
        console.warn(`No data found for selected week: ${selectedWeek}`);
        return;
    }

    categories.forEach(category => {
        const budget = parseInt(rowData[`${category}_Budget`] || 0);
        const availability = parseInt(rowData[`${category}_Availability`] || 0);
        const ordersLastYear = parseInt(rowData[`${category}_Orders_Last_Year`] || 0);

        updateElement(`this-week-${category.toLowerCase()}-budget`, budget);
        updateElement(`this-week-${category.toLowerCase()}-availability`, availability);
        updateElement(`this-week-${category.toLowerCase()}-orders-last-year`, ordersLastYear);
    });
}


// Utility: Update a DOM element by ID with formatted value
function updateElement(id, value) {
    const element = document.getElementById(id);
    if (element) {
        element.textContent = value.toLocaleString();
    } else {
        console.error(`Element with ID "${id}" not found.`);
    }
}

// Clear Section 2 data if no previous week is available
function clearPreviousWeekData() {
    categories.forEach(category => {
        updateElement(`prev-${category.toLowerCase()}-budget`, "N/A");
        updateElement(`prev-${category.toLowerCase()}-actual`, "N/A");
        updateElement(`prev-${category.toLowerCase()}-orders-last-year`, "N/A");
        updateElement(`prev-${category.toLowerCase()}-orders-this-year`, "N/A");
    });
}
