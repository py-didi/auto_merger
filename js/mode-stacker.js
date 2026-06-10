const ModeStacker = {
    triggerSheetScan(inputDom) {
        if(!inputDom.files[0]) return;
        ExcelEngine.getSheetNames(inputDom.files[0], sheets => {
            const container = document.getElementById("template-sheets-container");
            const checkboxesDiv = document.getElementById("template-sheets-checkboxes");
            checkboxesDiv.innerHTML = "";
            if(sheets.length > 0) {
                sheets.forEach((sheet, idx) => {
                    const checked = idx === 0 ? "checked" : "";
                    checkboxesDiv.innerHTML += `
                        <label style="display: inline-flex; align-items: center; gap: 4px; font-size: 11px; cursor: pointer; background: white; border: 1px solid #dfdfdf; padding: 2px 6px;">
                            <input type="checkbox" value="${sheet}" ${checked} style="margin: 0;">
                            <span style="color: black !important;">${sheet}</span>
                        </label>`;
                });
                container.classList.remove("hidden");
            }
        });
    },

    processTemplates() {
        const inputDom = document.getElementById("template-files-input");
        const commonHeaderIndex = parseInt(document.getElementById("template-header-idx").value) || 1;
        const checkedNodes = document.querySelectorAll("#template-sheets-checkboxes input:checked");
        const selectedSheets = Array.from(checkedNodes).map(n => n.value);

        if (!inputDom.files || inputDom.files.length === 0) {
            alert("Загрузите группу файлов шаблонов!");
            return;
        }

        App.showLoader("Извлечение заголовков первого шаблона...");
        setTimeout(() => {
            ExcelEngine.parseHeaders(inputDom.files[0], commonHeaderIndex, selectedSheets, headers => {
                App.hideLoader();
                if (!headers) return;

                const checkboxesContainer = document.getElementById("template-cols-checkboxes");
                checkboxesContainer.innerHTML = ""; // Никаких классов Tailwind

                headers.forEach(header => {
                    const labelWrapper = document.createElement("label");
                    // Жесткий инлайн-стиль для аккуратного вертикального списка
                    labelWrapper.style.display = "flex";
                    labelWrapper.style.alignItems = "center";
                    labelWrapper.style.gap = "6px";
                    labelWrapper.style.fontSize = "11px";
                    labelWrapper.style.cursor = "pointer";
                    labelWrapper.style.padding = "2px 4px";
                    labelWrapper.style.borderBottom = "1px solid #dfdfdf";
                    labelWrapper.style.width = "100%";
                    
                    labelWrapper.innerHTML = `<input type="checkbox" value="${header}" checked style="margin: 0;"> <span style="color: black; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${header}">${header}</span>`;
                    checkboxesContainer.appendChild(labelWrapper);
                });
                document.getElementById("template-mapping-area").classList.remove("hidden");
            });
        }, 50);
    },

    executeMerge() {
        const inputDom = document.getElementById("template-files-input");
        const commonHeaderIndex = parseInt(document.getElementById("template-header-idx").value) || 1;
        const includeFilenameOption = document.getElementById("template-add-filename").checked;

        const checkedNodes = document.querySelectorAll("#template-cols-checkboxes input:checked");
        const selectedTargetColumns = Array.from(checkedNodes).map(node => node.value);
        
        const sheetNodes = document.querySelectorAll("#template-sheets-checkboxes input:checked");
        const selectedSheets = Array.from(sheetNodes).map(n => n.value);

        if (selectedTargetColumns.length === 0) {
            alert("Выберите хотя бы один столбец для сбора данных!");
            return;
        }

        const totalFilesCount = inputDom.files.length;
        
        let csvChunks = ["\uFEFF"]; // BOM для Excel
        const headerArray = [...selectedTargetColumns];
        if (includeFilenameOption) headerArray.push("Источник_ИмяФайла");
        csvChunks.push(headerArray.map(ExcelEngine.escapeCSV).join(";") + "\n");

        App.showLoader("Начинаем пакетную сборку по частям...");
        
        setTimeout(() => {
            const processChunk = (index) => {
                if (index >= totalFilesCount) {
                    App.hideLoader();
                    if (csvChunks.length <= 1) { 
                        alert("Нет данных для объединения.");
                        return;
                    }
                    // ВЫЗОВ МОДАЛКИ СКАЧИВАНИЯ
                    ExcelEngine.promptDownload(csvChunks, "merged_templates_output");
                    return;
                }

                const currentFileObject = inputDom.files[index];
                document.getElementById("loading-text").innerText = `Обработка файла ${index + 1} из ${totalFilesCount}...`;
                
                setTimeout(() => {
                    const reader = new FileReader();
                    reader.onload = function(e) {
                        try {
                            let data = new Uint8Array(e.target.result);
                            let workbook = XLSX.read(data, { type: 'array', cellDates: false, cellNF: false, cellStyles: false });
                            data = null; 

                            const sheetsToProcess = selectedSheets && selectedSheets.length > 0 ? selectedSheets : [workbook.SheetNames[0]];
                            let chunkString = ""; 

                            sheetsToProcess.forEach(sheetName => {
                                if (!workbook.Sheets[sheetName]) return;
                                const worksheet = workbook.Sheets[sheetName];
                                
                                let rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "", blankrows: true });
                                const targetHeaderIdx = commonHeaderIndex - 1;
                                
                                if (rawRows.length <= targetHeaderIdx) {
                                    rawRows = null;
                                    return;
                                }
                                
                                const headerRowData = rawRows[targetHeaderIdx] || [];
                                const colIndexMap = {}; 
                                
                                let currentSheetHeaders = [];
                                for (let c = 0; c < headerRowData.length; c++) {
                                    let cellValue = headerRowData[c] ? String(headerRowData[c]).trim() : "";
                                    if (cellValue === "") cellValue = `Unnamed_${c + 1}`;
                                    
                                    let finalName = cellValue;
                                    let dupCounter = 1;
                                    while(currentSheetHeaders.includes(finalName)) {
                                        finalName = `${cellValue}_${dupCounter}`;
                                        dupCounter++;
                                    }
                                    currentSheetHeaders.push(finalName);
                                    
                                    if (selectedTargetColumns.includes(finalName)) {
                                        colIndexMap[finalName] = c;
                                    }
                                }
                                
                                for (let i = targetHeaderIdx + 1; i < rawRows.length; i++) {
                                    const currentRowData = rawRows[i];
                                    if (!currentRowData || currentRowData.length === 0) continue;
                                    
                                    const isEmptyRow = currentRowData.every(val => val === "" || val === null || val === undefined);
                                    if (isEmptyRow) continue;

                                    const rowArray = new Array(selectedTargetColumns.length);
                                    selectedTargetColumns.forEach((targetCol, tIdx) => {
                                        const colIdx = colIndexMap[targetCol];
                                        rowArray[tIdx] = colIdx !== undefined ? currentRowData[colIdx] : "";
                                    });
                                    
                                    if (includeFilenameOption) {
                                        rowArray.push(currentFileObject.name);
                                    }
                                    
                                    chunkString += rowArray.map(ExcelEngine.escapeCSV).join(";") + "\n";
                                }
                                rawRows = null; 
                            });
                            
                            workbook = null; 
                            csvChunks.push(chunkString); 
                            chunkString = null;

                        } catch (err) {
                            console.error(`Ошибка обработки файла ${currentFileObject.name}:`, err);
                        }
                        
                        setTimeout(() => processChunk(index + 1), 100);
                    };
                    reader.readAsArrayBuffer(currentFileObject);
                }, 50);
            };

            processChunk(0);
        }, 100);
    }
};