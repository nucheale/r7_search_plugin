(function (window, undefined) {
    window.Asc.plugin.init = async function () {
        let searchRangeChange = document.getElementById('search_range_change')
        const searchRangeValue = document.getElementById('search_range_value')
        let searchRangeSubmit = document.getElementById('search_range_submit')
        searchRangeChange.addEventListener('click', function() {
            console.info('change click handled')
            searchRangeValue.readOnly = false
        })
        searchRangeSubmit.addEventListener('click', async function () {
            let rangeValue = searchRangeValue.value
            // console.log(rangeValue)
            searchRangeValue.readOnly = true
        })

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

        let inputs = document.querySelectorAll('#sheet-name, #search-value');
        inputs.forEach(input => { //обновление resultMessage при изменении инпутов
            input.addEventListener('input', () => {
                if (resultMessage.innerText) resultMessage.innerText = '';
            });
        });
        document.getElementById('start-button').addEventListener('click', async function () {
            console.info('start click handled')
            resultMessage.innerText = 'Поиск...'
            await new Promise(resolve => setTimeout(resolve, 500)); // Пауза 0,5 сек для корректного обновления resultMessage
            let searchRange = searchRangeValue.value;
            let searchMode = document.querySelector('input[name="search-mode"]:checked').value;
            let sheetName = sheetSelect.value;
            let searchValue = document.getElementById('search-value').value;
            const result = await main(searchRange, sheetName, searchValue, searchMode);
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
        const searchRange = Asc.scope.searchRange
        if (!checkRange(searchRange)) {
            result = 'Неверный диапазон'
            return result
        }

        const searchValue = Asc.scope.searchValue;
        const searchMode = Asc.scope.searchMode;
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
        console.log(sheets);
        // конец список листов
        
        const resultArray = sheets.map(sheet => {
            const lastRow = getLastRow(sheet);
            const range = sheet.GetRange(`A1:ZZ${lastRow}`);
            return findValue(sheet, range, searchValue);
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

        function checkRange(range) {
            console.log(range)
            const regexRange = /^([A-Z]{1,3})(\d{1,7}):([A-Z]{1,3})(\d{1,7})$/i
            console.log(regexRange.test(range))
            if (!regexRange.test(range)) return false
            const [_, col1, row1, col2, row2] = regexRange.exec(range)
            console.log(col1 <= col2 && row1 <= row2)
            console.log([col1, col2, row1, row2])
            return col1 <= col2 && parseInt(row1) <= parseInt(row2)
        }

        function getLastRow(sh) {
            for (let row = 5000; row > 0; row--) {
                let rowValues = [];
                for (let col = 0; col < 701; col++) { // A-ZZ
                    let cellValue = sh.GetRangeByNumber(row, col).GetValue();
                    if (cellValue) rowValues.push(cellValue);
                }
                if (rowValues.join('').trim() !== '') {
                    console.log(`lastRow: ${row + 1}`)
                    return row + 1;
                }
            }
            return 5000
        }

        function findValue(sheet, range, value) {
            let findedCells = [];
            let data = range.GetValue();
            // value = String(value)
            data.forEach((row, rowIndex) => {
                row.forEach((cell, colIndex) => {
                    // cell = String(cell)
                    if (cell === value) {
                        findedCells.push(sheet.GetRangeByNumber(rowIndex, colIndex).GetAddress(false, false, "xlA1", false));
                    }
                });
            });

            return findedCells.length > 0? [sheet.GetName(), findedCells] : false
        }

    }

    async function main0() {
        return new Promise((resolve) => {
            window.Asc.plugin.callCommand(findSheetNames, false, false, function (value) {
                resolve(value)
            })
        })
    }

    async function main(searchRange, sheetName, searchValue, searchMode) {
        return new Promise((resolve) => {
            Asc.scope.searchRange = searchRange;
            Asc.scope.sheetName = sheetName;
            Asc.scope.searchValue = searchValue;
            Asc.scope.searchMode = searchMode;
            window.Asc.plugin.callCommand(findValuesInWorkbook, false, true, function (value) {
                resolve(value);
            });
        });
    }

})(window, undefined);
