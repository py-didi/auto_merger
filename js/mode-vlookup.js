const ModeVlookup = {
    mainFileConfig: null,
    extraFilesRegistry: [],
    extraBlocksCounter: 0,

    triggerSheetScanMain(inputDom) {
        if(!inputDom.files[0]) return;
        ExcelEngine.getSheetNames(inputDom.files[0], sheets => {
            const container = document.getElementById("main-sheets-container");
            const checkboxesDiv = document.getElementById("main-sheets-checkboxes");
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

    processMainFile() {
        const fileInput = document.getElementById("main-file-input");
        const headerRowNumber = parseInt(document.getElementById("main-header-idx").value) || 1;
        const checkedNodes = document.querySelectorAll("#main-sheets-checkboxes input:checked");
        const selectedSheets = Array.from(checkedNodes).map(n => n.value);

        if (!fileInput.files[0]) {
            alert("Выберите базовый файл таблицы!");
            return;
        }

        App.showLoader("Извлечение заголовков главного файла...");
        setTimeout(() => {
            ExcelEngine.parseHeaders(fileInput.files[0], headerRowNumber, selectedSheets, headers => {
                App.hideLoader();
                if (!headers) return;
                
                this.mainFileConfig = {
                    file: fileInput.files[0],
                    headerRow: headerRowNumber,
                    sheets: selectedSheets
                };

                const keySelector = document.getElementById("main-key-select");
                const checkboxesContainer = document.getElementById("main-cols-checkboxes");
                keySelector.innerHTML = "";
                checkboxesContainer.innerHTML = "";

                headers.forEach(header => {
                    const option = document.createElement("option");
                    option.value = header;
                    option.textContent = header;
                    keySelector.appendChild(option);

                    const labelWrapper = document.createElement("label");
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

                document.getElementById("main-mapping-area").classList.remove("hidden");
                document.getElementById("mode1-options").classList.remove("hidden");
                document.getElementById("mode1-controls").classList.remove("hidden");
                
                document.getElementById("extra-files-container").innerHTML = "";
                this.extraFilesRegistry = [];
            });
        }, 50);
    },

    addExtraFileBlock() {
        this.extraBlocksCounter++;
        const id = this.extraBlocksCounter;
        
        const container = document.getElementById("extra-files-container");
        const blockHTML = document.createElement("div");
        blockHTML.id = `extra-block-${id}`;
        blockHTML.style.border = "1px solid #808080";
        blockHTML.style.borderRightColor = "white";
        blockHTML.style.borderBottomColor = "white";
        blockHTML.style.padding = "12px";
        blockHTML.style.position = "relative";
        blockHTML.style.marginTop = "16px";
        
        blockHTML.innerHTML = `
            <div style="display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 8px;">
                <h3 style="font-weight: bold; color: #000080; margin:0;">⚙️ Доп. файл ${id}</h3>
                <button onclick="ModeVlookup.removeExtraFileBlock(${id})" class="win95-btn" style="padding: 2px 6px; font-size: 10px;">Удалить блок</button>
            </div>
            
            <div class="grid grid-cols-1 md:grid-cols-3 gap-4 items-end mb-3">
                <div class="md:col-span-2">
                    <input type="file" id="extra-file-${id}" accept=".xlsx, .xls, .csv" onchange="ModeVlookup.triggerSheetScanExtra(this, ${id})" class="win95-file-input">
                </div>
                <div>
                    <label style="display:block; font-size: 11px; margin-bottom:2px; color: black;">Строка шапки:</label>
                    <input type="number" id="extra-header-row-${id}" value="1" min="1" class="win95-input w-full text-center">
                </div>
            </div>

            <div id="extra-sheets-container-${id}" class="hidden win95-inset-pane mb-3">
                <label style="display:block; font-weight:bold; font-size: 11px; margin-bottom:4px; color: black;">Листы (FLAT склейка):</label>
                <div id="extra-sheets-checkboxes-${id}" class="flex flex-wrap gap-2"></div>
            </div>

            <button onclick="ModeVlookup.triggerExtraFileRead(${id})" class="win95-btn text-[11px]" style="margin-bottom: 8px;">
                Прочитать структуру доп. файла
            </button>

            <div id="extra-selectors-${id}" class="hidden grid grid-cols-1 md:grid-cols-2 gap-6 mt-2 pt-2 border-t border-gray-400">
                <div>
                    <label style="display:block; font-weight:bold; margin-bottom:4px; color: black;">Ключ связи в этой таблице:</label>
                    <select id="extra-key-select-${id}" class="win95-select w-full"></select>
                </div>
                <div>
                    <div style="display: flex; justify-content: space-between; align-items: flex-end; margin-bottom: 4px;">
                        <label style="font-weight:bold; color: black;">Какие столбцы подтянуть?</label>
                        <div style="display: flex; gap: 4px;">
                            <button onclick="App.toggleCheckboxes('extra-cols-checkboxes-${id}', true)" class="win95-btn" style="padding: 1px 6px; font-size: 10px;">Всё</button>
                            <button onclick="App.toggleCheckboxes('extra-cols-checkboxes-${id}', false)" class="win95-btn" style="padding: 1px 6px; font-size: 10px;">Снять</button>
                        </div>
                    </div>
                    <div id="extra-cols-checkboxes-${id}" class="win95-inset-pane" style="height: 100px; overflow-y: scroll; background-color: white;"></div>
                </div>
            </div>
        `;
        container.appendChild(blockHTML);
        this.extraFilesRegistry.push({ id, isActive: true });
    },

    removeExtraFileBlock(id) {
        document.getElementById(`extra-block-${id}`)?.remove();
        const match = this.extraFilesRegistry.find(item => item.id === id);
        if (match) match.isActive = false;
    },

    triggerSheetScanExtra(inputDom, id) {
        if(!inputDom.files[0]) return;
        ExcelEngine.getSheetNames(inputDom.files[0], sheets => {
            const container = document.getElementById(`extra-sheets-container-${id}`);
            const checkboxesDiv = document.getElementById(`extra-sheets-checkboxes-${id}`);
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

    triggerExtraFileRead(id) {
        const fileDom = document.getElementById(`extra-file-${id}`);
        const headerRowNumber = parseInt(document.getElementById(`extra-header-row-${id}`).value) || 1;
        const checkedNodes = document.querySelectorAll(`#extra-sheets-checkboxes-${id} input:checked`);
        const selectedSheets = Array.from(checkedNodes).map(n => n.value);

        if (!fileDom.files[0]) {
            alert("Сначала загрузите файл!");
            return;
        }

        App.showLoader("Извлечение заголовков доп. файла...");
        setTimeout(() => {
            ExcelEngine.parseHeaders(fileDom.files[0], headerRowNumber, selectedSheets, headers => {
                App.hideLoader();
                if (!headers) return;

                const target = this.extraFilesRegistry.find(item => item.id === id);
                if (target) {
                    target.fileConfig = {
                        file: fileDom.files[0],
                        headerRow: headerRowNumber,
                        sheets: selectedSheets
                    };
                }

                const keySelector = document.getElementById(`extra-key-select-${id}`);
                const checkboxesContainer = document.getElementById(`extra-cols-checkboxes-${id}`);
                keySelector.innerHTML = "";
                checkboxesContainer.innerHTML = "";

                headers.forEach(header => {
                    const option = document.createElement("option");
                    option.value = header;
                    option.textContent = header;
                    keySelector.appendChild(option);

                    const labelWrapper = document.createElement("label");
                    labelWrapper.style.display = "flex";
                    labelWrapper.style.alignItems = "center";
                    labelWrapper.style.gap = "6px";
                    labelWrapper.style.fontSize = "11px";
                    labelWrapper.style.cursor = "pointer";
                    labelWrapper.style.padding = "2px 4px";
                    labelWrapper.style.borderBottom = "1px solid #dfdfdf";
                    labelWrapper.style.width = "100%";
                    
                    labelWrapper.innerHTML = `<input type="checkbox" value="${header}" style="margin: 0;"> <span style="color: black; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;" title="${header}">${header}</span>`;
                    checkboxesContainer.appendChild(labelWrapper);
                });
                document.getElementById(`extra-selectors-${id}`).classList.remove("hidden");
            });
        }, 50);
    },

    executeMerge() {
        if (!this.mainFileConfig) return;
        
        App.showLoader("Чтение данных главного файла...");
        
        setTimeout(() => {
            ExcelEngine.parse(this.mainFileConfig.file, this.mainFileConfig.headerRow, this.mainFileConfig.sheets, mainOutput => {
                if (!mainOutput) { App.hideLoader(); return; }
                
                let loadedMainRows = mainOutput.rows;
                let loadedExtras = [];
                const activeExtras = this.extraFilesRegistry.filter(e => e.isActive && e.fileConfig);
                let currentExtraIdx = 0;

                const loadNextExtra = () => {
                    if (currentExtraIdx >= activeExtras.length) {
                        this._finalizeMerge(loadedMainRows, loadedExtras);
                        return;
                    }

                    const extraCfg = activeExtras[currentExtraIdx];
                    App.showLoader(`Чтение доп. файла ${currentExtraIdx + 1} из ${activeExtras.length}...`);
                    
                    setTimeout(() => {
                        ExcelEngine.parse(extraCfg.fileConfig.file, extraCfg.fileConfig.headerRow, extraCfg.fileConfig.sheets, extraOutput => {
                            if (extraOutput) {
                                loadedExtras.push({ id: extraCfg.id, rows: extraOutput.rows });
                            }
                            currentExtraIdx++;
                            loadNextExtra();
                        });
                    }, 50);
                };

                loadNextExtra();
            });
        }, 50);
    },

    _finalizeMerge(mainRows, loadedExtras) {
        App.showLoader("Проверка на дубликаты...");
        
        setTimeout(() => {
            const masterPrimaryKey = document.getElementById("main-key-select").value;
            let mainDupes = ExcelEngine.getDuplicatesCount(mainRows, masterPrimaryKey);
            let extraDupes = 0;
            
            loadedExtras.forEach(extra => {
                extra.foreignKey = document.getElementById(`extra-key-select-${extra.id}`).value;
                extraDupes += ExcelEngine.getDuplicatesCount(extra.rows, extra.foreignKey);
            });

            const autoRemoveDupes = document.getElementById("remove-dupes-checkbox").checked;
            let finalRemoveDupesFlag = autoRemoveDupes;

            if (!autoRemoveDupes && (mainDupes > 0 || extraDupes > 0)) {
                finalRemoveDupesFlag = confirm(`Внимание! Найдены дубликаты по ключевым столбцам.\n(В главном файле: ${mainDupes} шт., в дополнительных: ${extraDupes} шт.)\n\nУдалить дубликаты сейчас? \n[ОК] - Оставить только первые совпадения\n[Отмена] - Игнорировать предупреждение`);
            }

            App.showLoader("Подготовка к слиянию...");
            
            setTimeout(() => {
                const mainCheckedNodes = document.querySelectorAll("#main-cols-checkboxes input:checked");
                const masterFieldsToKeep = Array.from(mainCheckedNodes).map(node => node.value);
                if (!masterFieldsToKeep.includes(masterPrimaryKey)) masterFieldsToKeep.unshift(masterPrimaryKey);

                const activeExtraProcessors = [];
                loadedExtras.forEach(extra => {
                    const explicitCheckedNodes = document.querySelectorAll(`#extra-cols-checkboxes-${extra.id} input:checked`);
                    const fieldsToExtract = Array.from(explicitCheckedNodes).map(node => node.value);
                    
                    if (fieldsToExtract.length === 0) return;

                    let processedExtraRows = finalRemoveDupesFlag ? ExcelEngine.deduplicateRows(extra.rows, extra.foreignKey) : extra.rows;
                    const indexerLookupMap = new Map();
                    
                    processedExtraRows.forEach(record => {
                        const rawKeyValue = String(record[extra.foreignKey]).trim();
                        if (rawKeyValue && !indexerLookupMap.has(rawKeyValue)) {
                            indexerLookupMap.set(rawKeyValue, record);
                        }
                    });

                    activeExtraProcessors.push({ blockId: extra.id, fieldsToExtract, lookupMap: indexerLookupMap });
                });

                const dropUnmatchedRows = document.getElementById("drop-unmatched-rows").checked;
                let processedMainRows = finalRemoveDupesFlag ? ExcelEngine.deduplicateRows(mainRows, masterPrimaryKey) : mainRows;

                let csvChunks = ["\uFEFF"];
                const finalHeaders = [...masterFieldsToKeep];
                
                activeExtraProcessors.forEach(extra => {
                    extra.fieldsToExtract.forEach(field => {
                        let targetFieldName = masterFieldsToKeep.includes(field) ? `${field}_extra_${extra.blockId}` : field;
                        finalHeaders.push(targetFieldName);
                    });
                });
                
                csvChunks.push(finalHeaders.map(ExcelEngine.escapeCSV).join(";") + "\n");

                const totalRows = processedMainRows.length;
                let currentIndex = 0;
                const CHUNK_SIZE = 5000;

                const processMergeChunk = () => {
                    let chunkString = "";
                    const end = Math.min(currentIndex + CHUNK_SIZE, totalRows);

                    for (let i = currentIndex; i < end; i++) {
                        const masterRecord = processedMainRows[i];
                        const masterKeyValue = String(masterRecord[masterPrimaryKey]).trim();

                        let hasAnyMatch = false;
                        let rowArray = new Array(finalHeaders.length).fill("");

                        masterFieldsToKeep.forEach((field, idx) => {
                            rowArray[idx] = masterRecord[field] !== undefined ? masterRecord[field] : "";
                        });

                        let currentOffset = masterFieldsToKeep.length;
                        activeExtraProcessors.forEach(extra => {
                            const foundMatchedRecord = extra.lookupMap.get(masterKeyValue);
                            if (foundMatchedRecord) hasAnyMatch = true;

                            extra.fieldsToExtract.forEach((field, fIdx) => {
                                rowArray[currentOffset + fIdx] = foundMatchedRecord && foundMatchedRecord[field] !== undefined ? foundMatchedRecord[field] : "";
                            });
                            currentOffset += extra.fieldsToExtract.length;
                        });

                        if (dropUnmatchedRows) {
                            if (hasAnyMatch) {
                                chunkString += rowArray.map(ExcelEngine.escapeCSV).join(";") + "\n";
                            }
                        } else {
                            chunkString += rowArray.map(ExcelEngine.escapeCSV).join(";") + "\n";
                        }
                    }

                    csvChunks.push(chunkString);
                    chunkString = null;
                    currentIndex = end;

                    if (currentIndex < totalRows) {
                        const percent = Math.round((currentIndex / totalRows) * 100);
                        document.getElementById("loading-text").innerText = `Слияние таблиц: ${percent}%...`;
                        setTimeout(processMergeChunk, 15); 
                    } else {
                        App.hideLoader();
                        if (csvChunks.length <= 1) { 
                            alert("После применения всех фильтров не осталось ни одной строки!\nВозможно, ни одно значение из главного файла не было найдено в дополнительных.");
                            return;
                        }
                        
                        // ВЫЗОВ МОДАЛКИ СКАЧИВАНИЯ
                        ExcelEngine.promptDownload(csvChunks, "turbo_vlookup_output");
                        
                        loadedMainRows = null;
                        loadedExtras = null;
                    }
                };

                processMergeChunk();
            }, 50);
        }, 50);
    }
};