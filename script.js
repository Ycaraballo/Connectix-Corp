let processedBlob;
let config = 'PC-460';

document.getElementById('configSelector').addEventListener('change', (event) => {
    config = event.target.value;
    console.log('Config changed to:', config);
});

function processExcel() {
    const fileInput = document.getElementById('upload');
    const file = fileInput.files[0];

    if (!file) {
        alert("Please select a file first.");
        return;
    }

    const reader = new FileReader();

    reader.onload = function (e) {
        console.log('File loaded successfully.');
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet);

        console.log('Data parsed successfully:', jsonData);

        let processedData;
        
        if (config === 'PC-460') {
            processedData = processPC460(jsonData);
        } else {
            processedData = processPC430And600(jsonData);
        }

        console.log('Processed data:', processedData);

        // Mostrar los datos procesados en la tabla
        displayDataInTable(processedData);

        // Crear un nuevo libro de trabajo con los datos procesados
        const newSheet = XLSX.utils.json_to_sheet(processedData);
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Sheet1');

        // Generar archivo Excel y guardar el blob para descarga
        const wbout = XLSX.write(newWorkbook, { bookType: 'xlsx', type: 'array' });
        processedBlob = new Blob([wbout], { type: 'application/octet-stream' });

        // Mostrar el botón de descarga
        document.getElementById('downloadButton').style.display = 'block';
    };

    reader.onerror = function (error) {
        console.error('Error reading file:', error);
    };

    reader.readAsArrayBuffer(file);
}

function processPC460(data) {
    const priceMapping460 = {
        'CBI': 23.65,
        'CCoSu': 70.83,
        'CDI': 33.11,
        'CDP': 90.11,
        'CFDP': 107.59,
        'CFSP': 78.93,
        'CFTP': 147.44,
        'CRTC': 48.27,
        'CRTCDP': 52.38,
        'CRTCSP': 48.27,
        'CRTCTP': 60.34,
        'CSP': 66.11,
        'CTC': 37.84,
        'CTP': 123.48,
        'CWF': 21.28,
        'CWIFI': 28.38,
        'DB001': 28,
        'DB002': 38,
        'DB002A': 48,
        'RBI': 20.22,
        'RCoSu': 52.07,
        'RDI': 28.31,
        'RDP': 63.4,
        'RFDP': 75.7,
        'RFSP': 60.77,
        'RFTP': 86.72,
        'RRTCDP': 46.3,
        'RRTCSP': 42.67,
        'RRTCTP': 53.34,
        'RSP': 50.89,
        'RTC': 32.36,
        'RTP': 72.63,
        'RWF': 18.2
    };

    let sumGross = 0;
    let sumPaid = 0;

    const processedData = data.map(row => {
        // Sumar los valores originales en "Total" para Gross
        if (row['Total']) {
            let originalTotal = row['Total'];
            if (typeof originalTotal === 'number') {
                originalTotal = originalTotal.toString();
            }
            if (typeof originalTotal === 'string') {
                originalTotal = parseFloat(originalTotal.replace('$', ''));
                sumGross += originalTotal;
            }
        }

        // Calcular los valores en "Paid" usando los valores fijos multiplicados por QTY
        if (row['Job Code'] in priceMapping460) {
            const qty = row['QTY'] || 1;
            const fixedTotal = priceMapping460[row['Job Code']] * qty;
            sumPaid += fixedTotal;
            row['Total'] = `$${fixedTotal.toFixed(2)}`; // Actualizar la columna "Total" con el valor calculado
        }

        // Eliminar columnas no deseadas
        delete row['Office'];
        delete row['Work Area'];
        delete row['Emp ID'];
        delete row['Job #'];
        delete row['Project Name'];
        delete row['Work Type'];
        delete row['Rate']; // No mostrar Rate en PC-460

        return row;
    });

    // Actualizar los valores en el footer
    updateFooter(sumGross, sumPaid);

    return processedData;
}

function processPC430And600(data) {
    let sumGross = 0;
    let sumPaid = 0;

    const processedData = data.map(row => {
        if (row['Rate']) {
            let originalRate = row['Rate'];
            if (typeof originalRate === 'number') {
                originalRate = originalRate.toString();
            }
            if (typeof originalRate === 'string') {
                originalRate = parseFloat(originalRate.replace('$', ''));
                const qty = row['QTY'] || 1;
                const grossValue = originalRate * qty;
                const rate70 = originalRate * 0.7;
                const totalValue = rate70 * qty;

                row['Total'] = `$${totalValue.toFixed(2)}`;
                row['Rate'] = `$${rate70.toFixed(2)}`;

                sumGross += grossValue;
                sumPaid += totalValue;
            }
        }

        // Eliminar columnas no deseadas
        delete row['Office'];
        delete row['Work Area'];
        delete row['Emp ID'];
        delete row['Job #'];
        delete row['Project Name'];
        delete row['Work Type'];

        return row;
    });

    // Actualizar los valores en el footer
    updateFooter(sumGross, sumPaid);

    return processedData;
}

function displayDataInTable(data) {
    const tableHead = document.getElementById('tableHead');
    const tableBody = document.getElementById('tableBody');

    // Limpiar tabla antes de llenarla
    tableHead.innerHTML = '';
    tableBody.innerHTML = '';

    // Añadir encabezados
    const headers = Object.keys(data[0]);
    const headerRow = document.createElement('tr');
    headers.forEach(header => {
        const th = document.createElement('th');
        th.innerText = header;
        th.style.padding = '5px';  // Ajustar el tamaño del padding
        headerRow.appendChild(th);
    });
    tableHead.appendChild(headerRow);

    // Añadir filas de datos
    data.forEach(row => {
        const dataRow = document.createElement('tr');
        headers.forEach(header => {
            const td = document.createElement('td');
            td.innerText = row[header];
            td.style.padding = '5px';  // Ajustar el tamaño del padding
            dataRow.appendChild(td);
        });
        tableBody.appendChild(dataRow);
    });

    console.log('Table displayed successfully.');
}

function updateFooter(sumGross, sumPaid) {
    const difference = sumGross - sumPaid;
    document.getElementById('sumRate').innerText = `$${sumGross.toFixed(2)}`;
    document.getElementById('sumTotal').innerText = `$${sumPaid.toFixed(2)}`;
    document.getElementById('difference').innerText = `$${difference.toFixed(2)}`;
}

function downloadProcessedFile() {
    const url = URL.createObjectURL(processedBlob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'processed_file.xlsx';
    a.click();
    URL.revokeObjectURL(url); // Liberar memoria
}
