const ExcelEngine = {
    getSheetNames(fileObj, callback) {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', bookSheets: true });
                callback(workbook.SheetNames);
            } catch(err) {
                callback([]);
            }
        };
        reader.readAsArrayBuffer(fileObj);
    },

    parseHeaders(fileObj, humanHeaderRow, selectedSheets, callback) {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array', sheetRows: humanHeaderRow });
                
                let allHeaders = new Set();
                const sheetsToProcess = selectedSheets && selectedSheets.length > 0 ? selectedSheets : [workbook.SheetNames[0]];

                sheetsToProcess.forEach(sheetName => {
                    if (!workbook.Sheets[sheetName]) return;
                    const worksheet = workbook.Sheets[sheetName];
                    const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", blankrows: true });
                    const targetHeaderIdx = humanHeaderRow - 1;
                    
                    if (rawRows.length <= targetHeaderIdx) return; 
                    
                    let maxColumns = 0;
                    rawRows.forEach(row => { if (row.length > maxColumns) maxColumns = row.length; });

                    const headerRowData = rawRows[targetHeaderIdx] || [];
                    const currentSheetHeaders = [];
                    
                    for (let c = 0; c < maxColumns; c++) {
                        let cellValue = headerRowData[c] ? String(headerRowData[c]).trim() : "";
                        if (cellValue === "") cellValue = `Unnamed_${c + 1}`;
                        
                        let finalName = cellValue;
                        let dupCounter = 1;
                        while(currentSheetHeaders.includes(finalName)) {
                            finalName = `${cellValue}_${dupCounter}`;
                            dupCounter++;
                        }
                        currentSheetHeaders.push(finalName);
                        allHeaders.add(finalName); 
                    }
                });

                callback(Array.from(allHeaders));
            } catch (err) {
                alert(`Ошибка парсинга заголовков "${fileObj.name}": ` + err.message);
                callback(null);
            }
        };
        reader.readAsArrayBuffer(fileObj);
    },

    parse(fileObj, humanHeaderRow, selectedSheets, callback) {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                
                let allHeaders = new Set();
                let allRows = [];
                
                const sheetsToProcess = selectedSheets && selectedSheets.length > 0 ? selectedSheets : [workbook.SheetNames[0]];

                sheetsToProcess.forEach(sheetName => {
                    if (!workbook.Sheets[sheetName]) return;
                    const worksheet = workbook.Sheets[sheetName];
                    
                    const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", blankrows: true });
                    const targetHeaderIdx = humanHeaderRow - 1;
                    
                    if (rawRows.length <= targetHeaderIdx) return; 
                    
                    let maxColumns = 0;
                    rawRows.forEach(row => { if (row.length > maxColumns) maxColumns = row.length; });

                    const headerRowData = rawRows[targetHeaderIdx] || [];
                    const currentSheetHeaders = [];
                    
                    for (let c = 0; c < maxColumns; c++) {
                        let cellValue = headerRowData[c] ? String(headerRowData[c]).trim() : "";
                        if (cellValue === "") cellValue = `Unnamed_${c + 1}`;
                        
                        let finalName = cellValue;
                        let dupCounter = 1;
                        while(currentSheetHeaders.includes(finalName)) {
                            finalName = `${cellValue}_${dupCounter}`;
                            dupCounter++;
                        }
                        currentSheetHeaders.push(finalName);
                        allHeaders.add(finalName); 
                    }
                    
                    for (let i = targetHeaderIdx + 1; i < rawRows.length; i++) {
                        const currentRowData = rawRows[i];
                        if (!currentRowData || currentRowData.length === 0) continue;
                        
                        const isEmptyRow = currentRowData.every(val => val === "" || val === null || val === undefined);
                        if (isEmptyRow) continue;

                        const recordObject = {};
                        currentSheetHeaders.forEach((headerName, columnIdx) => {
                            let value = currentRowData[columnIdx];
                            recordObject[headerName] = (value !== undefined && value !== null) ? value : "";
                        });
                        allRows.push(recordObject);
                    }
                });

                callback({ headers: Array.from(allHeaders), rows: allRows });
            } catch (err) {
                alert(`Ошибка чтения файла "${fileObj.name}": ` + err.message);
                callback(null);
            }
        };
        reader.readAsArrayBuffer(fileObj);
    },

    getDuplicatesCount(rows, keyColumn) {
        let seen = new Set();
        let dupes = 0;
        for (let r of rows) {
            let val = String(r[keyColumn]).trim();
            if (val === "") continue; 
            if (seen.has(val)) dupes++;
            else seen.add(val);
        }
        return dupes;
    },

    deduplicateRows(rows, keyColumn) {
        let seen = new Set();
        return rows.filter(r => {
            let val = String(r[keyColumn]).trim();
            if (val === "") return true;
            if (seen.has(val)) return false;
            seen.add(val);
            return true;
        });
    },

    escapeCSV(val) {
        if (val === null || val === undefined) return "";
        let str = String(val);
        if (str.includes(';') || str.includes('"') || str.includes('\n') || str.includes('\r')) {
            str = '"' + str.replace(/"/g, '""') + '"';
        }
        return str;
    },

    exportCSV(chunksArray, filename) {
        const blob = new Blob(chunksArray, { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement("a");
        const url = URL.createObjectURL(blob);
        link.setAttribute("href", url);
        link.setAttribute("download", filename);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    },

    // --- НОВОЕ: Окно выбора скачивания (CSV или XLSX) ---
    promptDownload(csvChunks, baseFilename) {
        const overlay = document.createElement("div");
        overlay.className = "win95-overlay-bg flex";
        overlay.style.zIndex = "200";
        overlay.innerHTML = `
            <div class="win95-modal-window" style="width: 400px; padding: 2px;">
                <div class="title-bar">
                    <span class="title-bar-text pl-1">Сохранение файла</span>
                    <button class="win95-btn-title" onclick="this.closest('.win95-overlay-bg').remove()">X</button>
                </div>
                <div style="padding: 16px; background: var(--surface);">
                    <div style="display: flex; gap: 12px; align-items: center; margin-bottom: 20px;">
                        <span style="font-size: 32px;">💾</span>
                        <div>
                            <h3 style="font-weight: bold; font-size: 14px; margin:0 0 4px 0; color: #000080;">Сборка завершена!</h3>
                            <p style="margin:0; font-size: 11px;">Выберите формат итогового файла.</p>
                        </div>
                    </div>
                    <div style="display: flex; flex-direction: column; gap: 8px;">
                        <button id="btn-dl-csv" class="win95-btn" style="text-align: left; padding: 8px;">
                            📄 <strong style="font-size: 12px;">Скачать .CSV</strong><br>
                            <span style="font-weight:normal; font-size:10px; color: #333;">(Мгновенно. Рекомендуется для огромных таблиц)</span>
                        </button>
                        <button id="btn-dl-xlsx" class="win95-btn" style="text-align: left; padding: 8px;">
                            📊 <strong style="font-size: 12px;">Конвертировать в .XLSX (Excel)</strong><br>
                            <span style="font-weight:normal; font-size:10px; color: #333;">(Удобно. Но может вылететь от нехватки памяти ПК)</span>
                        </button>
                    </div>
                </div>
            </div>
        `;
        document.body.appendChild(overlay);

        document.getElementById("btn-dl-csv").onclick = () => {
            this.exportCSV(csvChunks, baseFilename + ".csv");
            overlay.remove();
        };

        document.getElementById("btn-dl-xlsx").onclick = () => {
            overlay.remove();
            App.showLoader("Конвертируем в Excel... Пожалуйста, подождите.");
            setTimeout(() => {
                try {
                    // Конвертация сырого CSV текста обратно в книгу Excel (тяжело для ОЗУ)
                    const csvString = csvChunks.join("");
                    const workbook = XLSX.read(csvString, {type: 'string', raw: true});
                    XLSX.writeFile(workbook, baseFilename + ".xlsx");
                    App.hideLoader();
                } catch (e) {
                    App.hideLoader();
                    alert("⚠️ Не хватило оперативной памяти для создания Excel-архива!\nСлишком большой объем данных.\n\nФайл будет сохранен в безопасном формате .CSV");
                    this.exportCSV(csvChunks, baseFilename + ".csv");
                }
            }, 100);
        };
    }
};