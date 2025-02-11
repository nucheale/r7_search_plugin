(function (window, undefined) {
    window.Asc.plugin.init = async function () {
        let resultMessage = document.getElementById('result-message');
        let sheetSelect = document.getElementById('sheet-name');
        sheetSelect.innerHTML = '';
        allSheets = [];
        allSheets = await main0();
        
        // setTimeout(() => {
        //     console.log("1s");
        // }, 1000);
        
        allSheets.forEach(sheet => {
            let option = document.createElement('option');
            option.value = sheet;
            option.textContent = sheet;
            if (sheet === "Отчетность") {
                option.selected = true;
            }
            sheetSelect.appendChild(option);
        });


        let inputs = document.querySelectorAll('#sheet-name, #search-value');
        inputs.forEach(input => { //обновление resultMessage при изменении инпутов
            input.addEventListener('input', () => {
                if (resultMessage.innerText === 'Ничего не найдено') {
                    resultMessage.innerText = '';
                }
            });
        });
        document.getElementById('start-button').addEventListener('click', async function () {
            console.info('start click handled')
            resultMessage.innerText = 'Поиск...'
            await new Promise(resolve => setTimeout(resolve, 500)); // Пауза 0,5 сек для корректного обновления resultMessage
            let searchMode = document.querySelector('input[name="search-mode"]:checked').value;
            let sheetName = document.getElementById('sheet-name').value;
            let searchValue = document.getElementById('search-value').value;
            const result = await main(sheetName, searchValue, searchMode);
            resultMessage.innerText = result;
        });
    };

    window.Asc.plugin.button = function (id) {
        this.executeCommand("close", "");
    };

    function findValuesInWorkbook() {
        const sheetsList = Api.GetSheets()
        console.log(sheetsList)

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
            let resultMessageCC = [];
            let findedCells = [];
            let data = range.GetValue();

            data.forEach((row, rowIndex) => {
                row.forEach((cell, colIndex) => {
                    if (cell === value) {
                        findedCells.push(sheet.GetRangeByNumber(rowIndex, colIndex).GetAddress(false, false, "xlA1", false));
                    }
                });
            });

            if (findedCells.length > 0) {
                console.info([findedCells])
                resultMessageCC.push(sheet.GetName())
                resultMessageCC.push([findedCells]);
            }
            return resultMessageCC.length > 0 ? resultMessageCC : false;
        }

        let sheetName;
        try {
            sheetName = Asc.scope.sheetName;
        } catch (error) {
        }
        const searchValue = Asc.scope.searchValue;
        const searchMode = Asc.scope.searchMode;

        let sh;
        let sheets = [];

        // получаем список листов
        switch (searchMode) {
            case 'single':
                try {
                    sh = Api.GetSheet(sheetName);
                    if (!sh) throw new Error(`Лист ${sheetName} не найден`)
                    sheets.push(sh)
                } catch (error) {
                    return error.message;
                }
                break;

            case 'all':
                const shs = Api.GetSheets();
                shs.forEach(sh => {
                    sheets.push(sh);
                });
                break;
        }
        // конец список листов

        let resultArray = [];
        sheets.forEach((sheet, i) => {
            console.info(sheet.GetName());
            const lastRow = getLastRow(sheet);
            const range = sheet.GetRange(`A1:ZZ${lastRow}`);
            const findedCells = findValue(sheet, range, searchValue);
            if (findedCells) resultArray.push(findedCells);
        });

        console.info(`resultArray: ${resultArray}`);
        let result = '';

        console.info(resultArray)

        if (resultArray.length > 0) {
            result += 'Найденные ячейки:\n\n'
            resultArray.forEach(element => {
                if (element.length > 0) {
                    result += `Лист '${element[0]}': ${element[1][0].join(', ')}\n\n`;
                }
            });

        } else {
            result = 'Ничего не найдено'
        }

        return result;

        // return result != ''
        //     ? `Ячейки на листе '${sheetName}' равные значению '${searchValue}': ${findedCells.join(', ')}`
        //     : `Ячейки на листе '${sheetName}' равные искомому значению '${searchValue}' не найдены`;
    }


    function findShNames() {
        let sheets = Api.GetSheets(); // Получаем все листы
        let result0 = []
        sheets.forEach(sheet => {
            result0.push(sheet.GetName())
        });
        return result0;
    }

    async function main0() {
        return new Promise((resolve) => {
            window.Asc.plugin.callCommand(findShNames, false, true, function (value) {
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
