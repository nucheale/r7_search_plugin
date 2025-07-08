(function (window, undefined) {
    window.Asc.plugin.init = async function () {
        let resultMessageSearch = document.getElementById('result-message-search');
        let resultMessageShowHideSheets = document.getElementById('result-message-show-hid-sheets');
        let resultMessageGetAllDefNames = document.getElementById('result-message-get-all-def-names');
        let resultMessageGetRefDefNames = document.getElementById('result-message-get-ref-def-names');
        let resultMessageShowDropdowns = document.getElementById('result-message-get-dropdowns');
        let sheetSelect = document.getElementById('sheet-name');
        let sheetCheckboxes = document.getElementById('sheet-names-checkboxes')
        sheetSelect.innerHTML = '';
        const [allSheets, activeSheet, workbookName] = await findSheetNamesInit();

        allSheets.forEach(sheet => {
            //добавляем в таб1 для выпадающего списка листов
            let option = document.createElement('option');
            option.value = option.textContent = sheet;
            if (sheet === activeSheet) option.selected = true;
            sheetSelect.appendChild(option);

            //добавляем в таб2 чекбоксы с названиями листов
            let label = document.createElement('label')
            let checkbox = document.createElement('input');
            checkbox.type = 'checkbox'
            checkbox.name = 'checkbox-show-hid-sheets'
            checkbox.value = sheet
            sheetCheckboxes.appendChild(label)
            label.appendChild(checkbox)
            label.appendChild(document.createTextNode(sheet))
        });

        //Убрать "Выбрать все" в чекбоксах. в markup не работает, т.к. чекбоксы генерируются здесь
        document.querySelectorAll('input[name="checkbox-show-hid-sheets"]:not([value="all"])').forEach((cb) => {
            cb.addEventListener("change", function () {
                const allChecked = Array.from(document.querySelectorAll('input[name="checkbox-show-hid-sheets"]:not([value="all"])'))
                    .every(cb => cb.checked);
                allCheckbox.checked = allChecked;
            });
        });

        const allInputs = document.querySelectorAll("input");
        allInputs.forEach((input) => {
            //обновление resultMessageSearch при изменении инпутов
            input.addEventListener("input", () => {
                if (resultMessageSearch.innerText) resultMessageSearch.innerText = "";
            });
        });

        const checkboxInputs = document.querySelectorAll('input[type="checkbox"]');
        checkboxInputs.forEach((input) => {
            //обновление resultMessageShowHideSheets при изменении инпутов
            input.addEventListener("input", () => {
                if (resultMessageShowHideSheets.innerText) resultMessageShowHideSheets.innerText = "";
            });
        });

        //клик "Найти"
        document.getElementById('start-button').addEventListener('click', async function () {
            console.info('start search click handled')
            resultMessageSearch.innerText = 'Поиск...'
            await new Promise(resolve => setTimeout(resolve, 500)); // Пауза 0,5 сек для корректного обновления resultMessageSearch

            // по типу поиска определяем значения для поиска (ЗНАЧ, ИМЯ, .., собственное значение) 
            const searchType = document.querySelector('input[name="search-type"]:checked').value;
            let searchValue
            if (searchType === 'any-text') {
                searchValue = document.getElementById('search-value').value
            } else {
                searchValue = searchType
            }

            // определяем остальные основные переменные
            const searchRange = document.getElementById('search_range_value').value
            const searchMode = document.querySelector('input[name="search-mode"]:checked').value;
            const searchMatch = document.querySelector('input[name="search-match"]:checked').value;
            const searchArea = document.querySelector('input[name="search-area"]:checked').value;
            const sheetName = sheetSelect.value;
            const resultSearch = await mainInitSearch(searchRange, sheetName, searchValue, searchMode, searchMatch, searchArea);
            resultMessageSearch.innerText = resultSearch;
        });

        //клики "Раскрыть" / "Скрыть"
        ['show-sheets-button', 'hide-sheets-button'].forEach(buttonId => {
            document.getElementById(buttonId).addEventListener('click', async function () {
                console.info('start hide-show-sheets click handled')
                // определяем листы для раскрытия
                let sheetsToShow = []
                sheetsToShow = Array.from(document.querySelectorAll('input[name="checkbox-show-hid-sheets"]:checked')).map(checked_sheet => checked_sheet.value);
                sheetsToShow = sheetsToShow.filter(value => value !== 'all')
                let resultShowSheets
                if (sheetsToShow.length > 0) {
                    resultShowSheets = await mainInitShowHideSheets(sheetsToShow, buttonId)
                    document.querySelectorAll('input[name="checkbox-show-hid-sheets"]').forEach(cb => { cb.checked = false }) //снимаем все чекбоксы после выполнения
                } else {
                    resultShowSheets = 'Не выбрано ни одного листа'
                }
                resultMessageShowHideSheets.innerText = resultShowSheets;
            });
        });

        //Клик "Отобразить" (именованные диапазоны на листе)
        // document.getElementById('get-def-names-button').addEventListener('click', async function () {
        //     console.log('start get-def-names-button handled')
        //     resultMessageGetDefNames = await mainGetDefNamesOnSheet()
        // })

        //Выбор папки для всех именованных диапазонов на листе
        document.getElementById("get-all-def-names-button").addEventListener("change", event => {
            const files = event.target.files
            let matchedFile
            for (const file of files) {
                const filePath = file.webkitRelativePath
                if (filePath.includes(`/${workbookName}`)) {
                    matchedFile = file
                    break
                } else {
                    continue
                }
            }
            console.log(matchedFile)
            if (!matchedFile) return

            const reader = new FileReader();
            reader.onload = function (e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: "array" })

                let defNames
                if (workbook.Workbook.Names) {
                    defNames = workbook.Workbook.Names
                } else {
                    defNames = []
                }

                let activeSheetDefNames = []
                defNames.forEach(defName => {
                    const defNameRef = defName.Ref
                    defNameRef.includes(activeSheet) ? activeSheetDefNames.push([defName.Name, defName.Ref]) : false
                })

                let output
                if (defNames.length > 0) {
                    output = activeSheetDefNames.map(defName => {
                        return `${defName[0]}: ${defName[1]}`
                    })
                    output = output.join("\n\n")
                } else {
                    output = 'На активном листе отсутствуют именованные диапазоны'
                }

                resultMessageGetAllDefNames.innerText = output
            }

            reader.readAsArrayBuffer(matchedFile)
        })


        //Выбор папки для REF именованных диапазонов в КНИГЕ
        document.getElementById("get-ref-def-names-button").addEventListener("change", event => {
            const files = event.target.files
            let matchedFile
            for (const file of files) {
                const filePath = file.webkitRelativePath
                if (filePath.includes(`/${workbookName}`)) {
                    matchedFile = file
                    break
                } else {
                    continue
                }
            }

            if (!matchedFile) {
                resultMessageGetRefDefNames.innerText = 'Папка не содержит активный файл'
                return
            }

            const reader = new FileReader();
            reader.onload = function (e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: "array" })

                let defNames
                if (workbook.Workbook.Names) {
                    defNames = workbook.Workbook.Names
                } else {
                    defNames = []
                }

                let refDefNames = []
                defNames.forEach(defName => {
                    const defNameRef = defName.Ref
                    defNameRef.includes('#REF') ? refDefNames.push([defName.Name, defName.Ref]) : false
                })

                let output
                if (defNames.length > 0) {
                    output = refDefNames.map(defName => {
                        return `${defName[0]}: ${defName[1]}`
                    })
                    output = output.join("\n\n")
                } else {
                    output = 'REF диапазоны не найдены'
                }

                resultMessageGetRefDefNames.innerText = output

            }

            reader.readAsArrayBuffer(matchedFile)
        })

    }




    window.Asc.plugin.button = function (id) {
        this.executeCommand("close", "");
    };


    function findSheetNames() {
        const workbookName = Api.GetFullName()
        const activeSheet = Api.GetActiveSheet().GetName()
        const sheets = Api.GetSheets();
        let allSheets = []
        sheets.forEach(sheet => {
            allSheets.push(sheet.GetName())
        });
        return [allSheets, activeSheet, workbookName];
    };

    function findValuesInWorkbook() {
        let result = '';

        // проверяем корректность пользовательского диапазона для поиска
        let searchRange = Asc.scope.searchRange
        if (searchRange == '') searchRange = 'A1:XFD1048576'
        if (!checkRange(searchRange)) {
            result = 'Неверный диапазон'
            return result
        }

        const searchValue = Asc.scope.searchValue;
        const searchMode = Asc.scope.searchMode;
        const searchMatch = Asc.scope.searchMatch;
        const searchArea = Asc.scope.searchArea;
        const sheetName = Asc.scope.sheetName;
        let sheets = [];
        // получаем список листов для поиска
        switch (searchMode) {
            case 'single':
                sheets = [Api.GetSheet(sheetName)]
                break;
            case 'all':
                sheets = Api.GetSheets();
                break;
        }
        // console.log(sheets);
        // конец список листов

        // ищем, записываем в итоговый массив имя листа и адреса ячееек/ячейки
        const resultArray = sheets.map(sheet => {
            const lastRow = getLastRow(sheet, searchRange)
            if (lastRow === 0) return // если лист пустой - пропускаем
            const lastCol = columnNumberToLetter(getLastCol(sheet, searchRange))
            //lastCol добавить в функции!!!!!!!!!!!!!!!!!!!!!!!!! или что-то другое. посмотреть
            let sheetSearchRange = searchRange.replace(/\d+$/, lastRow) //замена последней строки
            sheetSearchRange = sheetSearchRange.replace(/:([A-Z]+)(\d+)/, `:${lastCol}$2`) //замена последнего столбца
            // const range = sheet.GetRange(sheetSearchRange)
            const { start: [startRow, startCol], end: [endRow, endCol] } = parsexlRange(sheetSearchRange)
            console.log(`sheet: ${sheet.GetName()}`)
            console.log(`lastRow: ${lastRow}`)
            // проверяем как искать - по значениям или по формулам, запускаем нужную функцию
            switch (searchArea) {
                case 'values':
                    return findValue(sheet, sheetSearchRange, searchValue, startRow, startCol, searchMatch);
                case 'formulas':
                    return findFormula(sheet, sheetSearchRange, searchValue, startRow, startCol, endRow, endCol, searchMatch)
                // return findFormulaApi(sheet, range, searchValue, endRow, searchMatch)
            }
        }).filter(Boolean);
        console.info('resultArray: ', resultArray)

        if (resultArray.length > 0) {
            result += 'Найденные ячейки:\n\n'
            resultArray.forEach(element => {
                if (element.length > 0) {
                    result += `Лист '${element[0]}': ${element[1].join(', ')}\n\n`
                }
            });

        } else {
            result = 'Ничего не найдено'
        }

        return result


        // фукнция проверки корректности введенного диапазона (соответствует формату диапазона Excel)
        function checkRange(range) {
            const regexRange = /^([A-Z]{1,3})(\d{1,7}):([A-Z]{1,3})(\d{1,7})$/i
            if (!regexRange.test(range)) return false
            const [_, col1, row1, col2, row2] = regexRange.exec(range)
            return col1 <= col2 && parseInt(row1) <= parseInt(row2) //проверяется корректность диапазона, например Z100:A10 не пройдет обе проверки
        }

        // преобразование адреса вида XL (A1:ZZ5000) в координаты (0, 0, 4999, 701)
        function parsexlRange(range) {
            const [startCell, endCell] = range.split(':')

            function parseXLCell(cell) {
                const cellRegex = /^([A-Za-z]+)(\d+)$/
                const match = cell.match(cellRegex)
                const colLetters = match[1].toUpperCase();
                const rowNumber = parseInt(match[2], 10);
                let colNumber = 0
                for (let i = 0; i < colLetters.length; i++) {
                    colNumber = colNumber * 26 + (colLetters.charCodeAt(i) - 65 + 1)
                }
                return {
                    col: colNumber - 1,
                    row: rowNumber - 1
                }
            }

            const start = parseXLCell(startCell)
            const end = parseXLCell(endCell)

            // console.log(start.row, start.col, end.row, end.col)

            return {
                start: [start.row, start.col],
                end: [end.row, end.col]
            }
        }

        //преобразования номера столбца в буквенное представление
        function columnNumberToLetter(columnNumber) {
            let letter = ''
            while (columnNumber > 0) {
                const remainder = (columnNumber - 1) % 26
                letter = String.fromCharCode(65 + remainder) + letter
                columnNumber = Math.floor((columnNumber - 1) / 26)
            }
            return letter
        }


        //функция нахождения последней строки в диапазоне
        function getLastRow(sh, xlRange) {
            const isFullSheet = (xlRange == 'A1:XFD1048576') ? true : false
            let data
            if (isFullSheet) {
                data = sh.GetUsedRange().GetValue()
            } else {
                data = sh.GetRange(xlRange).GetValue()
            }
            for (let i = data.length - 1; i >= 0; i--) {
                const isEmptyRow = data[i].every(cell => cell === undefined || cell === '')
                if (!isEmptyRow) {
                    if (isFullSheet) {
                        return i + 1
                    } else {
                        const { start } = parsexlRange(xlRange)
                        const startRow = start[0]
                        return i + startRow + 1
                    }
                }
            }
            return 0
        }

        function getLastCol(sh, xlRange) {
            const isFullSheet = (xlRange == 'A1:XFD1048576') ? true : false
            let data
            if (isFullSheet) {
                data = sh.GetUsedRange().GetValue()
            } else {
                data = sh.GetRange(xlRange).GetValue()
            }
            let lastCol = 0
            for (let i = 0; i < data.length; i++) {
                for (let j = data[i].length; j >= 0; j--) {
                    const cell = data[i][j]
                    if (cell !== undefined && cell !== '') {
                        if (j + 1 > lastCol) lastCol = j + 1
                        break
                    }
                }
            }

            if (!isFullSheet) {
                const { start } = parsexlRange(xlRange);
                const startCol = start[1];
                lastCol += startCol;
            }

            return lastCol;

        }



        // функция для поиска ячеек с нужным значением в определенном диапазоне
        function findValue(sheet, range, value, startRow, startCol, searchMatch, searchArea) {
            console.info('findValue started')
            let findedCells = []
            const data = sheet.GetRange(range).GetValue()
            const normalizedValue = value.toLowerCase()
            data.forEach((row, rowIndex) => {
                row.forEach((cell, colIndex) => {
                    const match = searchMatch === 'exact' ? cell.toLowerCase() === normalizedValue : cell.toLowerCase().includes(normalizedValue)
                    if (match) {
                        const address = sheet.GetRangeByNumber(rowIndex + startRow, colIndex + startCol).GetAddress(false, false, 'xlA1', false)
                        findedCells.push(address)
                    }
                });
            });

            return findedCells.length > 0 ? [sheet.GetName(), findedCells] : false
        }

        function findFormula(sheet, range, value, startRow, startCol, endRow, endCol, searchMatch) {
            console.info('findFormula started')
            let findedCells = []
            const normalizedValue = value.toLowerCase()

            for (let i = startRow; i <= endRow; i++) {
                for (let j = startCol; j <= endCol; j++) {
                    const cell = sheet.GetRangeByNumber(i, j).GetFormula()
                    const match = searchMatch === 'exact' ? cell.toLowerCase() === normalizedValue : cell.toLowerCase().includes(normalizedValue)
                    if (match) {
                        const address = sheet.GetRangeByNumber(i, j).GetAddress(false, false, 'xlA1', false)
                        findedCells.push(address)
                    }
                }
            }

            return findedCells.length > 0 ? [sheet.GetName(), findedCells] : false
        }

        // function findFormulaApi(sheet, range, value, endRow, searchMatch) {
        //     let findedCells = []
        //     let firstFoundedCell
        //     const matchType = searchMatch === 'exact' ? "xlWhole" : "xlPart"
        //     firstFoundedCell = range.Find(value, "A1", "xlFormulas", matchType, "xlByColumns", "xlNext", false)
        //     if (!firstFoundedCell) return false
        //     let firstAddress = sheet.GetRange(firstFoundedCell).GetAddress(false, false, "xlA1", false)
        //     findedCells.push(firstAddress)
        //     let currentCell = firstFoundedCell
        //     for (let i = 1; i <= endRow + 1; i++) {
        //         let nextCell = range.FindNext(currentCell)
        //         let nextAddress = sheet.GetRange(nextCell).GetAddress(false, false, "xlA1", false)
        //         if (nextAddress === firstAddress) break //выход из цикла если больше ничего не найдено (т.е. нашли ту же ячейку)
        //         findedCells.push(nextAddress)
        //         currentCell = nextCell
        //     }

        //     return findedCells.length > 0 ? [sheet.GetName(), findedCells] : false
        // }
    }

    function showHideSheets() {
        const sheets = Asc.scope.sheetNames
        let mode = Asc.scope.mode
        let visableState = false
        if (mode === 'show-sheets-button') visableState = true
        const activeSheet = Api.GetActiveSheet()

        sheets.forEach(sh => {
            let sheet = Api.GetSheet(sh)
            if (sheet.GetVisible()) sheet.SetActive() //для корректного скрытия листа
            if (sheet.GetName() !== activeSheet.GetName()) sheet.SetVisible(visableState) //с активным листом ничего не делаем
        });

        if (!sheets.includes(activeSheet)) activeSheet.SetActive()
        return 'Выполнено'
    }

    function getDefNamesOnSheet() {
        const sheet = Api.GetActiveSheet()
        const sheets = Api.GetSheets()
        const sfd = Api.GetFullName()
        console.log(sfd)
    }

    async function findSheetNamesInit() {
        return new Promise((resolve) => {
            window.Asc.plugin.callCommand(findSheetNames, false, false, function (value) {
                resolve(value)
            })
        })
    }

    async function mainInitSearch(searchRange, sheetName, searchValue, searchMode, searchMatch, searchArea) {
        return new Promise((resolve) => {
            if (searchValue === '') {
                resolve('Не введено значение для поиска');
                return
            }
            Asc.scope.searchRange = searchRange;
            Asc.scope.sheetName = sheetName;
            Asc.scope.searchValue = searchValue;
            Asc.scope.searchMode = searchMode;
            Asc.scope.searchMatch = searchMatch;
            Asc.scope.searchArea = searchArea;
            window.Asc.plugin.callCommand(findValuesInWorkbook, false, true, function (value) {
                resolve(value);
            });
        });
    }

    async function mainInitShowHideSheets(sheets, mode) {
        return new Promise((resolve) => {
            Asc.scope.sheetNames = sheets
            Asc.scope.mode = mode
            window.Asc.plugin.callCommand(showHideSheets, false, true, function (value) {
                resolve(value)
            })
        })
    }

    // async function mainGetDefNamesOnSheet() {
    //     return new Promise((resolve) => {
    //         window.Asc.plugin.callCommand(getDefNamesOnSheet, false, false, function (value) {
    //             resolve(value)
    //         })
    //     })
    // }

})(window, undefined);
