(function (window, undefined) {
    window.Asc.plugin.init = async function () {
        let resultMessage = document.getElementById('result-message');
        let sheetSelect = document.getElementById('sheet-name');
        sheetSelect.innerHTML = '';
        const [allSheets, activeSheet] = await main0();

        allSheets.forEach(sheet => {
            let option = document.createElement('option');
            option.value = option.textContent = sheet;
            if (sheet === activeSheet) option.selected = true;
            sheetSelect.appendChild(option);
        });

        const inputs = document.querySelectorAll('input');
        inputs.forEach(input => { //обновление resultMessage при изменении инпутов
            input.addEventListener('input', () => {
                if (resultMessage.innerText) resultMessage.innerText = '';
            });
        });
        document.getElementById('start-button').addEventListener('click', async function () {
            console.info('start click handled')
            resultMessage.innerText = 'Поиск...'
            await new Promise(resolve => setTimeout(resolve, 500)); // Пауза 0,5 сек для корректного обновления resultMessage

            // по типу поиска определяем значения для поиска (ЗНАЧ, ИМЯ, .., собственное значение) 
            const searchType = document.querySelector('input[name="search-type"]:checked').value;
            let searchValue
            if (searchType === 'any-text') {
                searchValue = document.getElementById('search-value').value;
            } else {
                searchValue = searchType
            }

            // определяем остальные основные переменные
            const searchRange = document.getElementById('search_range_value').value
            const searchMode = document.querySelector('input[name="search-mode"]:checked').value;
            const searchMatch = document.querySelector('input[name="search-match"]:checked').value;
            const searchArea = document.querySelector('input[name="search-area"]:checked').value;
            const sheetName = sheetSelect.value;
            const result = await main(searchRange, sheetName, searchValue, searchMode, searchMatch, searchArea);
            resultMessage.innerText = result;
        });
    };

    window.Asc.plugin.button = function (id) {
        this.executeCommand("close", "");
    };


    function findSheetNames() {
        const activeSheet = Api.GetActiveSheet().GetName()
        const sheets = Api.GetSheets();
        let allSheets = []
        sheets.forEach(sheet => {
            allSheets.push(sheet.GetName())
        });
        return [allSheets, activeSheet];
    };

    function findValuesInWorkbook() {
        let result = '';

        // проверяем корректность диапазона для поиска
        const searchRange = Asc.scope.searchRange
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
            const lastRow = getLastRow(sheet, searchRange);
            const sheetSearchRange = searchRange.replace(/\d+$/, lastRow)
            const range = sheet.GetRange(sheetSearchRange);
            const { start, end } = parseXLRange(sheetSearchRange) //исправить логику
            const startRow = start[0] //исправить логику
            const startCol = start[1] //исправить логику
            const endRow = end[0] //исправить логику
            const endCol = end[1] //исправить логику
            // проверяем как искать - по значениям или по формулам, запускаем нужную функцию
            switch (searchArea) {
                case 'values':
                    return findValue(sheet, range, searchValue, startRow, startCol, searchMatch);
                case 'formulas':
                    // return findFormula(sheet, searchValue, startRow, startCol, endRow, endCol, searchMatch)
                    return findFormulaApi(sheet, range, searchValue, endRow, searchMatch)
            }
        }).filter(Boolean);
        console.info('resultArray: ');
        console.info(resultArray);

        if (resultArray.length > 0) {
            result += 'Найденные ячейки:\n\n'
            resultArray.forEach(element => {
                if (element.length > 0) {
                    result += `Лист '${element[0]}': ${element[1].join(', ')}\n\n`;
                }
            });

        } else {
            result = 'Ничего не найдено'
        }

        return result;


        // фукнция проверки корректности введенного диапазона (соответствует формату диапазона Excel)
        function checkRange(range) {
            const regexRange = /^([A-Z]{1,3})(\d{1,7}):([A-Z]{1,3})(\d{1,7})$/i
            if (!regexRange.test(range)) return false
            const [_, col1, row1, col2, row2] = regexRange.exec(range)
            return col1 <= col2 && parseInt(row1) <= parseInt(row2) //проверяется корректность диапазона, например Z100:A10 не пройдет обе проверки
        }

        // преобразование адреса вида XL (A1:ZZ5000) в координаты (0, 0, 4999, 701)
        function parseXLRange(range) {
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


        // фукнция нахождения последней заполненной строки в диапазоне
        function getLastRow(sh, XLrange) {
            const { start, end } = parseXLRange(XLrange)
            const startRow = start[0]
            const startCol = start[1]
            const endRow = end[0]
            const endCol = end[1]

            for (let row = endRow; row >= startRow; row--) {
                let rowValues = [];
                for (let col = startCol; col <= endCol; col++) {
                    let cellValue = sh.GetRangeByNumber(row, col).GetValue();
                    if (cellValue) rowValues.push(cellValue);
                }
                if (rowValues.join('').trim() !== '') {
                    console.log(`lastRow: ${row + 1}`)
                    return row + 1;
                }
            }
            return endRow
        }

        // функция для поиска ячеек с нужным значением в определенном диапазоне
        function findValue(sheet, range, value, startRow, startCol, searchMatch, searchArea) {
            let findedCells = [];
            const data = range.GetValue();
            const normalizedValue = value.toLowerCase()
            data.forEach((row, rowIndex) => {
                row.forEach((cell, colIndex) => {
                    const match = searchMatch === 'exact' ? cell.toLowerCase() === normalizedValue : cell.toLowerCase().includes(normalizedValue);
                    if (match) {
                        const address = sheet.GetRangeByNumber(rowIndex + startRow, colIndex + startCol).GetAddress(false, false, "xlA1", false)
                        findedCells.push(address);
                    }
                });
            });

            return findedCells.length > 0 ? [sheet.GetName(), findedCells] : false
        }

        function findFormula(sheet, value, startRow, startCol, endRow, endCol, searchMatch) {
            let findedCells = [];
            const normalizedValue = value.toLowerCase()
            for (let i = startRow; i <= endRow; i++) {
                for (let j = startCol; j <= endCol; j++) {
                    let cellFormula = sheet.GetRangeByNumber(i, j).GetFormula()
                    cellFormula = cellFormula.slice(0, 1) + cellFormula.slice(2)
                    console.log(cellFormula.toLowerCase())
                    const match = searchMatch === 'exact' ? cellFormula.toLowerCase() === normalizedValue : cellFormula.toLowerCase().includes(normalizedValue);
                    if (match) {
                        const address = sheet.GetRangeByNumber(i, j).GetAddress(false, false, "xlA1", false)
                        findedCells.push(address);
                    }
                }
            }

            return findedCells.length > 0 ? [sheet.GetName(), findedCells] : false
        }

        function findFormulaApi(sheet, range, value, endRow, searchMatch) {
            let findedCells = [];
            let firstFoundedCell = range.Find(value, "A1", "xlFormulas", "xlPart", "xlByColumns", "xlNext", false);
            if (!firstFoundedCell) return false;
            let firstAddress = sheet.GetRange(firstFoundedCell).GetAddress(false, false, "xlA1", false);
            findedCells.push(firstAddress);
            let currentCell = firstFoundedCell;
            for (let i = 1; i <= endRow; i++) {
                let nextCell = range.FindNext(currentCell);
                let nextAddress = sheet.GetRange(nextCell).GetAddress(false, false, "xlA1", false);
                if (nextAddress === firstAddress) break //выход из цикла если больше ничего не найдено (т.е. нашли ту же ячейку)
                findedCells.push(nextAddress);
                currentCell = nextCell;
            }

            return [sheet.GetName(), findedCells];
        }
    }

    async function main0() {
        return new Promise((resolve) => {
            window.Asc.plugin.callCommand(findSheetNames, false, false, function (value) {
                resolve(value)
            })
        })
    }

    async function main(searchRange, sheetName, searchValue, searchMode, searchMatch, searchArea) {
        return new Promise((resolve) => {
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

})(window, undefined);
