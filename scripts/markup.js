// Скрытие/отображение ввода имени листа в зависимости от выбранного режима
document
    .querySelectorAll('input[name="search-mode"]')
    .forEach((radio) => {
        radio.addEventListener("change", function () {
            const sheetNameContainer = document.getElementById("sheet-name-container")
            if (this.value === "single") {
                sheetNameContainer.classList.remove("hidden");
            } else {
                sheetNameContainer.classList.add("hidden");
            }
        });
    });

// Скрытие/отображение ввода текста поиска в зависимости от выбранного режима
document
    .querySelectorAll('input[name="search-type"]')
    .forEach((radio) => {
        radio.addEventListener("change", function () {
            const valueContainer = document.getElementById("value-container");
            if (this.value === "any-text") {
                valueContainer.classList.remove("hidden");
            } else {
                valueContainer.classList.add("hidden");
            }
        });
    });

// Вкладки
document.querySelectorAll(".tab-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
        const tabId = btn.dataset.tab;

        document
            .querySelectorAll(".tab-btn")
            .forEach((b) => b.classList.remove("active"));
        btn.classList.add("active");

        document.querySelectorAll(".tab").forEach((tab) => {
            tab.classList.remove("active");
            if (tab.id === tabId) {
                tab.classList.add("active");
            }
        });
    });
});


// Чекбокс "Выбрать все"
let allCheckbox = document.querySelector('input[name="checkbox-show-hid-sheets"][value="all"]')
allCheckbox.addEventListener("change", function () {
    document.querySelectorAll('input[name="checkbox-show-hid-sheets"]:not([value="all"])').forEach((cb) => {
        cb.checked = this.checked;
    });
});

