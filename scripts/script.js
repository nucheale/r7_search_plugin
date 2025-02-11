(function (window, undefined) {
    window.Asc.plugin.init = async function () {
        let resultMessage = document.getElementById('result-message');
        let sheetSelect = document.getElementById('sheet-name');
        sheetSelect.innerHTML = '';
        const allSheets = await main0();
        
        // setTimeout(() => {
        //     console.log("1s");
        // }, 1000);
        
        allSheets.forEach(sheet => {
            let option = document.createElement('option');
            option.value = option.textContent = sheet;
            if (sheet === "Отчетность") option.selected = true;
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
            let searchMode = document.querySelector('input[name="search-mode"]:checked').value;
            let sheetName = sheetSelect.value;
            let searchValue = document.getElementById('search-value').value;
            const result = await main(sheetName, searchValue, searchMode);
            resultMessage.innerText = result;
        });
    };

    window.Asc.plugin.button = function (id) {
        this.executeCommand("close", "");
    };

    function findShNames() {
        const sheets = Api.GetSheets();
        let result0 = []
        sheets.forEach(sheet => {
            result0.push(sheet.GetName())
        });
        return result0;
    };

    function findValuesInWorkbook() {
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

        let result = '';
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

        function getLastRow(sh) {
            for (let row = 1000; row > 0; row--) {
                let rowValues = [];
                for (let col = 0; col < 701; col++) { // A-ZZ
                    let cellValue = sh.GetRangeByNumber(row, col).GetValue();
                    if (cellValue) rowValues.push(cellValue);
                }
                if (rowValues.join('').trim() !== '') {
                    console.info(row)
                    return row;
                }
            }
            return 1000
        }

        function findValue(sheet, range, value) {
            let findedCells = [];
            let data = range.GetValue();

            data.forEach((row, rowIndex) => {
                row.forEach((cell, colIndex) => {
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
            window.Asc.plugin.callCommand(findShNames, false, false, function (value) {
                resolve(value)
            })
        })
    }

    async function main(sheetName, searchValue, searchMode) {
        return new Promise((resolve) => {
            Asc.scope.sheetName = sheetName;
            Asc.scope.searchValue = searchValue;
            Asc.scope.searchMode = searchMode;
            window.Asc.plugin.callCommand(findValuesInWorkbook, false, true, function (value) {
                resolve(value);
            });
        });
    }

})(window, undefined);
