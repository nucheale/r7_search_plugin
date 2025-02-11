(function (window, undefined) {
    window.Asc.plugin.init = function () {
        const sheetName = document.getElementById('sheet-name').value;
        let NBSPMessage = document.getElementById('nbsp-message');
        document.getElementById('start-button').addEventListener('click', async function () {
            const result = await main();
            NBSPMessage.innerText = result;
        });
    };

    window.Asc.plugin.button = function (id) {
        this.executeCommand("close", "");
    };

    function findNBSP() {
        function getLastRow(sh) {
            for (let i = 3; i < 50000; i++) {
                if (!sh.GetRangeByNumber(i, 0).GetValue()) {
                    return i;
                }
            }
        }

        function findValue(sheet, range, value) {
            let data = range.GetValue();
            let findedCells = [];

            data.forEach((row, rowIndex) => {
                row.forEach((cell, colIndex) => {
                    if (cell === value) {
                        findedCells.push(sheet.GetRangeByNumber(rowIndex, colIndex).GetAddress(false, false, "xlA1", false));
                    }
                });
            });
            return findedCells.length > 0 ? findedCells : [];
        }

        const sh = Api.GetSheet('Отчетность');
        const lastRow = getLastRow(sh);
        const range = sh.GetRange(`A1:Z${lastRow}`);

        const findedCells = findValue(sh, range, ' '); // Неразрывный пробел

        return findedCells.length > 0
            ? `Ячейки с неразрывным пробелом: ${findedCells.join(', ')}`
            : "Ячейки с неразрывным пробелом не найдены";
    }

    async function main() {
        return new Promise((resolve) => {
            window.Asc.plugin.callCommand(findNBSP, false, true, function (value) {
                resolve(value);
            });
        });
    }

})(window, undefined);
