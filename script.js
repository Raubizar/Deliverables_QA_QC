let fileData = [];

document.getElementById('fileInput').addEventListener('change', function (event) {
    const file = event.target.files[0];
    if (!file) {
        alert("Please select an Excel file.");
        return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        processExcelData(jsonData);
    };

    reader.readAsArrayBuffer(file);
});

function normalizeText(value) {
    if (typeof value === 'string') {
        return value.trim().replace(/\s+/g, ' ').toUpperCase();
    }
    return value;
}

function normalizeDate(value) {
  if (typeof value !== 'string') return value;

  // Try multiple formats
  const formats = [
    /^\d{2}\.\d{2}\.\d{4}$/,  // 13.03.2025
    /^\d{2}\/\d{2}\/\d{4}$/,  // 13/03/2025
    /^\d{2}\.\w{3}\.\d{4}$/i, // 13.MAR.2025
    /^\d{2}\/\d{2}\/\d{2}$/   // 13/03/25
  ];

  // First, see if input matches one of the patterns
  const isKnownFormat = formats.some(regex => regex.test(value.trim()));
  if (!isKnownFormat) {
    // Fallback to default JS Date parse
    const fallbackDate = new Date(value);
    if (!isNaN(fallbackDate.getTime())) {
      // Convert to dd/mm/yyyy for consistency
      return fallbackDate.toLocaleDateString('en-GB', {
        day: '2-digit',
        month: '2-digit',
        year: 'numeric'
      });
    }
    // Return original string if parsing fails
    return value.trim();
  }

  // If it does match, parse accordingly
  // (Simplest approach: replace "." with "/" then try new Date)
  const standardized = value.trim()
    .replace(/\./g, '/')  // Convert dots to slashes
    .replace(/([A-Za-z]{3})/i, (m) => {
      // Convert "MAR" to "03" etc. (This would need a small map)
      const monthMap = { JAN: '01', FEB: '02', MAR: '03', APR: '04', MAY: '05', JUN: '06', JUL: '07', AUG: '08', SEP: '09', OCT: '10', NOV: '11', DEC: '12' };
      return monthMap[m.toUpperCase()] || m;
    });

  const parsedDate = new Date(standardized);
  if (!isNaN(parsedDate.getTime())) {
    return parsedDate.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' });
  }

  return value.trim();
}

function processExcelData(data) {
    fileData = data.slice(1).map(row => ({
        sheetNumber: normalizeText(row[0]),
        sheetName: normalizeText(row[1]),
        fileName: normalizeText(row[2]),
        revisionCode: normalizeText(row[3]),
        revisionDate: normalizeDate(row[4]),
        suitabilityCode: normalizeText(row[5]),
        stageDescription: normalizeText(row[6]),
        documentNamingConvention: 'OK',
        comments: '',
        result: 'Pending',
        mismatches: ''
    }));

    populateTable();
}

function populateTable() {
    const tableBody = document.querySelector('#reportTable tbody');
    tableBody.innerHTML = '';

    fileData.forEach((row, index) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row.sheetNumber}</td>
            <td>${row.sheetName}</td>
            <td>${row.fileName}</td>
            <td>${row.revisionCode}</td>
            <td>${row.revisionDate}</td>
            <td>${row.suitabilityCode}</td>
            <td>${row.stageDescription}</td>
            <td>
                <select class="documentNamingConvention" data-index="${index}">
                    <option value="OK" selected>OK</option>
                    <option value="Not correct">Not correct</option>
                </select>
            </td>
            <td><input type="text" class="commentsInput" data-index="${index}" placeholder="Add comments"></td>
            <td id="result-${index}">${row.result}</td>
            <td id="mismatch-${index}">${row.mismatches}</td>
        `;
        tableBody.appendChild(tr);
    });

    document.getElementById('output-section').style.display = 'block';
}

document.getElementById('processFile').addEventListener('click', function () {
    document.querySelectorAll('.documentNamingConvention').forEach(select => {
        const index = select.getAttribute('data-index');
        fileData[index].documentNamingConvention = select.value;
    });

    document.querySelectorAll('.commentsInput').forEach(input => {
        const index = input.getAttribute('data-index');
        fileData[index].comments = normalizeText(input.value);
    });

    calculateResults();
});

function calculateResults() {
    let totalFiles = fileData.length;
    let okCount = 0;
    const expectedRevCode = normalizeText(document.getElementById('revisionCode').value || '');
    const expectedRevDate = normalizeDate(document.getElementById('revisionDate').value || '');
    const expectedSuitCode = normalizeText(document.getElementById('suitabilityCode').value || '');
    const expectedStageDesc = normalizeText(document.getElementById('stageDescription').value || '');
    const expectedRevisionDesc = document.getElementById('revisionDescription').value.trim();
    const separator = document.getElementById('separator').value || ' - ';  // Default separator if missing
    const checkOnlySheetNumber = document.getElementById('checkOnlySheetNumber').checked;

    fileData.forEach((row, index) => {
        let mismatches = [];

        // Handle missing fields by considering them as mismatches
        const sheetNumber = row.sheetNumber ? normalizeText(row.sheetNumber) : '';
        const sheetName = row.sheetName ? normalizeText(row.sheetName) : '';
        const fileName = row.fileName ? normalizeText(row.fileName) : '';
        const revisionCode = row.revisionCode ? normalizeText(row.revisionCode) : '';
        const revisionDate = row.revisionDate ? normalizeDate(row.revisionDate) : '';
        const suitabilityCode = row.suitabilityCode ? normalizeText(row.suitabilityCode) : '';
        const stageDescription = row.stageDescription ? normalizeText(row.stageDescription) : '';
        const documentNamingConvention = row.documentNamingConvention || 'Not correct';
        const comments = row.comments ? normalizeText(row.comments) : '';

        if (!sheetNumber || !sheetName || !fileName || !revisionCode || !revisionDate || !suitabilityCode || !stageDescription) {
            mismatches.push('Missing Data');
        }

        let nameCheck;
        if (checkOnlySheetNumber) {
            // Only compare sheetNumber to fileName
            nameCheck = (sheetNumber === fileName);
        } else {
            // Compare sheetNumber + separator + sheetName to fileName
            nameCheck = (sheetNumber + separator + sheetName) === fileName;
        }

        if (!nameCheck) mismatches.push('File Name');

        let revisionValid = revisionCode.startsWith(expectedRevCode[0]) && parseInt(revisionCode.slice(1)) >= parseInt(expectedRevCode.slice(1));
        if (!revisionValid) mismatches.push('Revision Code');

        let dateValid = revisionDate === expectedRevDate;
        if (!dateValid) mismatches.push('Revision Date');

        let suitabilityValid = suitabilityCode === expectedSuitCode;
        if (!suitabilityValid) mismatches.push('Suitability Code');

        let stageDescValid = stageDescription === expectedStageDesc;
        if (!stageDescValid) mismatches.push('Stage Description');

        let revisionDescValid = row.revisionDescription === expectedRevisionDesc;
        if (!revisionDescValid) mismatches.push('Revision Description');

        let namingConventionValid = documentNamingConvention === "OK";
        if (!namingConventionValid) mismatches.push('Document Naming Convention');

        let commentsValid = comments === '';
        if (!commentsValid) mismatches.push('Comments');

        let isValid = mismatches.length === 0;

        const resultCell = document.getElementById(`result-${index}`);
        const mismatchCell = document.getElementById(`mismatch-${index}`);

        if (isValid) {
            resultCell.textContent = "OK";
            fileData[index].result = "OK";
            okCount++;
        } else {
            resultCell.textContent = "Please Revise";
            fileData[index].result = "Please Revise";
        }
        
        mismatchCell.textContent = mismatches.join(', ');
        fileData[index].mismatches = mismatches.join(', ');
    });

    document.getElementById('totalFiles').textContent = totalFiles;
    document.getElementById('percentOK').textContent = ((okCount / totalFiles) * 100).toFixed(2) + '%';
    document.getElementById('summary-section').style.display = 'block';
}


document.getElementById('exportReport').addEventListener('click', function () {
    exportCombinedCSV();
});

function exportCombinedCSV() {
    let csvContent = "data:text/csv;charset=utf-8,";

    // Add summary at the top
    csvContent += "SUMMARY REPORT\n";
    csvContent += "Total Files,Percentage OK\n";
    csvContent += `${fileData.length},${((fileData.filter(row => row.result === "OK").length / fileData.length) * 100).toFixed(2)}%\n\n`;

    // Add full report header including mismatch column
    csvContent += "FULL REPORT\n";
    

    function safeValue(value) {
        if (value === null || value === undefined) {
            return '"MISSING"';
        }
        return `"${String(value).replace(/"/g, '""').trim()}"`;  // Convert to string, escape quotes, and trim safely
    }

    // Add CSV headers
    csvContent += '"Sheet Number","Sheet Name","File Name","Revision Code","Revision Date","Suitability Code","Stage Description","Document Naming Convention","Comments","Result","Mismatched Items"\n';

    // Process each row
    fileData.forEach(row => {
        csvContent += [
            safeValue(row.sheetNumber),
            safeValue(row.sheetName),
            safeValue(row.fileName),
            safeValue(row.revisionCode),
            safeValue(row.revisionDate),
            safeValue(row.suitabilityCode),
            safeValue(row.stageDescription),
            safeValue(row.documentNamingConvention),
            safeValue(row.comments || ""),
            safeValue(row.result),
            safeValue(row.mismatches || "NONE")
        ].join(",") + "\n";
    });

    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "QA_QC_Report.csv");
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}



