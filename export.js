// Логика экспорта Excel вынесена в отдельный файл.
// Файл использует глобальные данные и функции, объявленные в script.js.

function getExcelColumnName(index) {
    let dividend = index;
    let columnName = "";
    while (dividend > 0) {
        const modulo = (dividend - 1) % 26;
        columnName = String.fromCharCode(65 + modulo) + columnName;
        dividend = Math.floor((dividend - modulo) / 26);
    }
    return columnName;
}

function makeExcelBorder({ top = "thin", bottom = "thin", left = "thin", right = "thin" } = {}) {
    const color = { argb: "FF000000" };
    return {
        top: { style: top, color },
        bottom: { style: bottom, color },
        left: { style: left, color },
        right: { style: right, color }
    };
}

function styleExcelCell(cell, options = {}) {
    if (options.font) cell.font = options.font;
    if (options.fill) {
        cell.fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: options.fill }
        };
    }
    if (options.alignment) cell.alignment = options.alignment;
    if (options.border) cell.border = options.border;
}

function getExportDayCellAppearance(cell) {
    if (cell.kind === "empty") {
        return { topValue: "", bottomValue: "", fill: "FFFFFFFF" };
    }

    if (cell.absence) {
        if (cell.absence.type === "Отгул") {
            return { topValue: cell.code || getAbsenceCode(cell.absence.type), bottomValue: "", fill: "FFDCEEFF" };
        }
        if (cell.absence.type === "Срочный больничный") {
            return { topValue: cell.code || getAbsenceCode(cell.absence.type), bottomValue: "", fill: "FFFFF2A8" };
        }
        if (cell.absence.type === "Больничный") {
            return { topValue: cell.code || getAbsenceCode(cell.absence.type), bottomValue: "", fill: "FFFFF2A8" };
        }
        return { topValue: cell.code || getAbsenceCode(cell.absence.type), bottomValue: "", fill: "FF8EC9FF" };
    }

    const topValue = cell.worked ? (cell.hours || "") : "";
    const bottomValue = cell.worked ? (cell.night || "") : "";

    if (cell.holiday) {
        return { topValue, bottomValue, fill: "FFF7D9E6" };
    }

    if (cell.preholiday) {
        return { topValue, bottomValue, fill: "FFFFF5CC" };
    }

    if (cell.weekend) {
        return { topValue, bottomValue, fill: "FFF9DADA" };
    }

    if (cell.worked) {
        return {
            topValue,
            bottomValue,
            fill: "FFFFFFFF"
        };
    }

    return {
        topValue: "",
        bottomValue: "",
        fill: "FFFFFFFF"
    };
}


function addAnnualSummarySection(worksheet, startRow, graph) {
    const annualStats = buildAnnualStats(graph);
    const summaryStartCol = 1;
    const summaryEndCol = 9;
    const summaryTitleRow = startRow;
    const summaryHeaderRow = startRow + 1;
    let currentRow = startRow + 2;

    worksheet.mergeCells(summaryTitleRow, summaryStartCol, summaryTitleRow, summaryEndCol);
    const titleCell = worksheet.getCell(summaryTitleRow, summaryStartCol);
    titleCell.value = 'Итоговая сводка по сменам';
    styleExcelCell(titleCell, {
        font: { name: 'Times New Roman', size: 12, bold: true },
        alignment: { horizontal: 'center', vertical: 'middle' },
        fill: 'FFF2F2F2',
        border: makeExcelBorder({ top: 'medium', bottom: 'medium', left: 'medium', right: 'medium' })
    });
    worksheet.getRow(summaryTitleRow).height = 22;

    const headers = [
        'Смена',
        'Факт. дней',
        'Факт. часов',
        'Ночных',
        'Норма часов',
        'Часы отвлечений',
        'Плановые отвлечения, дн',
        'Кратковременные отвлечения, дн',
        'Баланс'
    ];

    headers.forEach((header, index) => {
        const cell = worksheet.getCell(summaryHeaderRow, summaryStartCol + index);
        cell.value = header;
        styleExcelCell(cell, {
            font: { name: 'Times New Roman', size: 11, bold: true },
            alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
            fill: 'FFF2F2F2',
            border: makeExcelBorder({
                top: 'medium',
                bottom: 'medium',
                left: index === 0 ? 'medium' : 'thin',
                right: index === headers.length - 1 ? 'medium' : 'thin'
            })
        });
    });
    worksheet.getRow(summaryHeaderRow).height = 34;

    annualStats.forEach((item, idx) => {
        const row = currentRow + idx;
        const balanceLabel = item.diffHours > 0
            ? `Переработка +${item.diffHours} ч`
            : item.diffHours < 0
                ? `Недоработка ${Math.abs(item.diffHours)} ч`
                : 'Норма выполнена';

        const values = [
            item.smena.name,
            item.workedDays,
            item.hours,
            item.night,
            item.normHours,
            item.absenceHours,
            item.plannedAbsenceDays,
            item.shortAbsenceDays,
            balanceLabel
        ];

        values.forEach((value, index) => {
            const cell = worksheet.getCell(row, summaryStartCol + index);
            cell.value = value;
            styleExcelCell(cell, {
                font: { name: 'Times New Roman', size: 11 },
                alignment: { horizontal: index === 0 || index === 8 ? 'left' : 'center', vertical: 'middle', wrapText: true },
                border: makeExcelBorder({
                    top: 'thin',
                    bottom: idx === annualStats.length - 1 ? 'medium' : 'thin',
                    left: index === 0 ? 'medium' : 'thin',
                    right: index === headers.length - 1 ? 'medium' : 'thin'
                })
            });
        });
        worksheet.getRow(row).height = 22;
    });

    worksheet.getColumn(1).width = 22;
    worksheet.getColumn(2).width = 12;
    worksheet.getColumn(3).width = 14;
    worksheet.getColumn(4).width = 12;
    worksheet.getColumn(5).width = 14;
    worksheet.getColumn(6).width = 16;
    worksheet.getColumn(7).width = 20;
    worksheet.getColumn(8).width = 24;
    worksheet.getColumn(9).width = 18;

    return currentRow + annualStats.length;
}

async function exportActiveGraphToExcel() {
    const graph = getActiveGraph();
    if (!graph) {
        alert("Нет активного графика для экспорта.");
        return;
    }
    if (!window.ExcelJS) {
        alert("Библиотека экспорта не загружена.");
        return;
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "ShiftMaster";
    workbook.created = new Date();

    const worksheet = workbook.addWorksheet("График", {
        views: [{ state: "frozen", xSplit: 2, ySplit: 11, showGridLines: false }]
    });

    worksheet.pageSetup = {
        orientation: "landscape",
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
        paperSize: 9,
        margins: {
            left: 0.3,
            right: 0.3,
            top: 0.35,
            bottom: 0.35,
            header: 0.2,
            footer: 0.2
        }
    };

    const dayStartCol = 3;
    const dayEndCol = 33;
    const graphDaysCol = 34;
    const graphHoursCol = 35;
    const prodDaysCol = 36;
    const prodHoursCol = 37;
    const corrHoursCol = 38;
    const absenceDaysCol = 39;

    worksheet.columns = [
        { width: 13 },
        { width: 10 },
        ...Array.from({ length: 31 }, () => ({ width: 5.2 })),
        { width: 7.5 },
        { width: 8.5 },
        { width: 8.5 },
        { width: 10.5 },
        { width: 16.8 },
        { width: 23.8 }
    ];

    const title = `График сменности сотрудников ${graph.name} при ${graph.type === "24" ? "24-часовой" : "12-часовой"} смене на ${currentYear} год`;
    worksheet.mergeCells(6, dayStartCol, 6, absenceDaysCol);
    const titleCell = worksheet.getCell(6, dayStartCol);
    titleCell.value = title;
    styleExcelCell(titleCell, {
        font: { name: "Times New Roman", size: 14, bold: true },
        alignment: { horizontal: "center", vertical: "middle" }
    });
    worksheet.getRow(6).height = 24;

    worksheet.mergeCells("A9:A11");
    worksheet.mergeCells("B9:B11");
    worksheet.mergeCells(`C9:${getExcelColumnName(dayEndCol)}10`);
    worksheet.mergeCells(`${getExcelColumnName(graphDaysCol)}9:${getExcelColumnName(graphHoursCol)}10`);
    worksheet.mergeCells(`${getExcelColumnName(prodDaysCol)}9:${getExcelColumnName(prodHoursCol)}10`);
    worksheet.mergeCells(`${getExcelColumnName(corrHoursCol)}9:${getExcelColumnName(corrHoursCol)}11`);
    worksheet.mergeCells(`${getExcelColumnName(absenceDaysCol)}9:${getExcelColumnName(absenceDaysCol)}11`);

    worksheet.getCell("A9").value = "месяц";
    worksheet.getCell("B9").value = "смена";
    worksheet.getCell("C9").value = "Часов в день, в том числе ночных";
    worksheet.getCell(`${getExcelColumnName(graphDaysCol)}9`).value = "По графику";
    worksheet.getCell(`${getExcelColumnName(prodDaysCol)}9`).value = "По произ. кален.";
    worksheet.getCell(`${getExcelColumnName(corrHoursCol)}9`).value = "Корректировка (час)";
    worksheet.getCell(`${getExcelColumnName(absenceDaysCol)}9`).value = "Дней отвлечений из рабочих";

    for (let day = 1; day <= 31; day += 1) {
        worksheet.getCell(11, dayStartCol + day - 1).value = day;
    }

    worksheet.getCell(11, graphDaysCol).value = "дней";
    worksheet.getCell(11, graphHoursCol).value = "часов";
    worksheet.getCell(11, prodDaysCol).value = "дней";
    worksheet.getCell(11, prodHoursCol).value = "часов";

    const headerFill = "FFF2F2F2";
    const headerAlignment = { horizontal: "center", vertical: "middle", wrapText: true };
    for (let row = 9; row <= 11; row += 1) {
        for (let col = 1; col <= absenceDaysCol; col += 1) {
            const cell = worksheet.getCell(row, col);
            styleExcelCell(cell, {
                font: { name: "Times New Roman", size: 11, bold: true },
                fill: headerFill,
                alignment: headerAlignment,
                border: makeExcelBorder({
                    top: row === 9 ? "medium" : "thin",
                    bottom: row === 11 ? "medium" : "thin",
                    left: col === 1 ? "medium" : "thin",
                    right: col === absenceDaysCol ? "medium" : "thin"
                })
            });
        }
    }

    let currentRow = 12;

    for (let monthIndex = 0; monthIndex < 12; monthIndex += 1) {
        const monthRows = graph.smeny.map((smena, smenaIndex) => buildMonthRowData(graph, smena, smenaIndex, monthIndex));
        const monthStartRow = currentRow;
        const monthEndRow = currentRow + monthRows.length * 2 - 1;

        worksheet.mergeCells(monthStartRow, 1, monthEndRow, 1);
        const monthCell = worksheet.getCell(monthStartRow, 1);
        monthCell.value = MONTHS[monthIndex].toLowerCase();
        styleExcelCell(monthCell, {
            font: { name: "Times New Roman", size: 11 },
            alignment: { horizontal: "center", vertical: "middle" },
            border: makeExcelBorder({
                top: "medium",
                bottom: "medium",
                left: "medium",
                right: "thin"
            })
        });

        for (const rowData of monthRows) {
            const topRow = currentRow;
            const bottomRow = currentRow + 1;

            worksheet.mergeCells(topRow, 2, bottomRow, 2);
            const shiftCell = worksheet.getCell(topRow, 2);
            shiftCell.value = rowData.smena.name.toLowerCase();
            styleExcelCell(shiftCell, {
                font: { name: "Times New Roman", size: 11 },
                alignment: { horizontal: "center", vertical: "middle" },
                border: makeExcelBorder({
                    top: "medium",
                    bottom: bottomRow === monthEndRow ? "medium" : "thin",
                    left: "thin",
                    right: "thin"
                })
            });

            rowData.cells.forEach((cell, index) => {
                const col = dayStartCol + index;
                const appearance = getExportDayCellAppearance(cell);
                const topCell = worksheet.getCell(topRow, col);
                const bottomCell = worksheet.getCell(bottomRow, col);

                topCell.value = appearance.topValue || null;
                bottomCell.value = appearance.bottomValue || null;

                styleExcelCell(topCell, {
                    font: { name: "Times New Roman", size: 11 },
                    fill: appearance.fill,
                    alignment: { horizontal: "center", vertical: "middle" },
                    border: makeExcelBorder({
                        top: "medium",
                        bottom: "thin",
                        left: "thin",
                        right: "thin"
                    })
                });
                styleExcelCell(bottomCell, {
                    font: { name: "Times New Roman", size: 11 },
                    fill: appearance.fill,
                    alignment: { horizontal: "center", vertical: "middle" },
                    border: makeExcelBorder({
                        top: "thin",
                        bottom: bottomRow === monthEndRow ? "medium" : "thin",
                        left: "thin",
                        right: "thin"
                    })
                });
            });

            const absenceWorkingDays = rowData.cells.filter((cell) =>
                cell.kind === "day" && cell.absence && !cell.weekend && !cell.holiday
            ).length;
            const correctionHours = absenceWorkingDays * 8;
            const correctedProductionHours = Math.max(0, rowData.productionStats.workHours - correctionHours);

            worksheet.getCell(topRow, graphDaysCol).value = rowData.rowStats.workedDays;
            worksheet.getCell(topRow, graphHoursCol).value = rowData.rowStats.hours;
            worksheet.getCell(topRow, prodDaysCol).value = rowData.productionStats.workDays;
            worksheet.getCell(topRow, prodHoursCol).value = correctedProductionHours;
            worksheet.getCell(topRow, corrHoursCol).value = correctionHours;
            worksheet.getCell(topRow, absenceDaysCol).value = absenceWorkingDays;

            for (const col of [graphDaysCol, graphHoursCol, prodDaysCol, prodHoursCol, corrHoursCol, absenceDaysCol]) {
                styleExcelCell(worksheet.getCell(topRow, col), {
                    font: { name: "Times New Roman", size: 11 },
                    alignment: { horizontal: "center", vertical: "middle" },
                    border: makeExcelBorder({
                        top: "medium",
                        bottom: "thin",
                        left: col === graphDaysCol ? "medium" : "thin",
                        right: col === absenceDaysCol ? "medium" : "thin"
                    })
                });
                styleExcelCell(worksheet.getCell(bottomRow, col), {
                    font: { name: "Times New Roman", size: 11 },
                    alignment: { horizontal: "center", vertical: "middle" },
                    border: makeExcelBorder({
                        top: "thin",
                        bottom: bottomRow === monthEndRow ? "medium" : "thin",
                        left: col === graphDaysCol ? "medium" : "thin",
                        right: col === absenceDaysCol ? "medium" : "thin"
                    })
                });
            }

            worksheet.getRow(topRow).height = 21;
            worksheet.getRow(bottomRow).height = 18;
            currentRow += 2;
        }
    }

    currentRow += 2;
    addAnnualSummarySection(worksheet, currentRow, graph);

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob(
        [buffer],
        { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
    );
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const fileName = `${graph.name.replace(/[^a-zа-я0-9_-]+/gi, "_")}_${currentYear}.xlsx`;

    link.href = url;
    link.download = fileName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
}

window.addEventListener("resize", handleResize);
