import * as XLSX from "xlsx-js-style";

interface ExportSheet {
  data: object[];
  sheetName?: string;
  calcName?: string;
  calcCol?: string;
  calcFormula?: string;
  calcResult?: number;
}

export const exportSheet = (exportData: ExportSheet | ExportSheet[]) => {
  const cellColor = {
    fill: { fgColor: { rgb: "1e81c6" } },
  };
  const headerColor = {
    font: { sz: 14, color: { rgb: "ffffff" }, bold: true },
  };
  const columnWidth = 20;
  const cellBorder = {
    border: {
      top: { style: "thin", color: { rgb: "00000" } },
      right: { style: "thin", color: { rgb: "00000" } },
      bottom: { style: "thin", color: { rgb: "00000" } },
      left: { style: "thin", color: { rgb: "00000" } },
    },
  };

  if (!Array.isArray(exportData)) {
    try {
      const { data, sheetName, calcName, calcCol, calcFormula, calcResult } =
        exportData;
      const dataKeys = Object.keys(data[0]);

      const workSheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(data);

      //Get the col and row range
      const range = XLSX.utils.decode_range(workSheet["!ref"] ?? "");
      const rowCount = range.e.r;
      const columnCount = range.e.c;

      //Format
      for (let row = 0; row <= rowCount; row++) {
        for (let col = 0; col <= columnCount; col++) {
          const cellRef = XLSX.utils.encode_cell({ r: row, c: col });

          // Format every cell
          workSheet[cellRef].s = {
            ...cellBorder,
          };
          if (col === 0) {
            //Format first column
            workSheet[cellRef].s = {
              alignment: { horizontal: "center" },
              ...cellBorder,
            };
          }
          if (row === 0) {
            // Format first row
            workSheet[cellRef].s = {
              ...workSheet[cellRef].s,
              ...cellColor,
              ...cellBorder,
              ...headerColor,
              alignment: { horizontal: "center" },
            };
          }
        }
      }

      const workSheetColumns = [];
      for (let i = 0; i < dataKeys.length; i++) {
        workSheetColumns.push({ wch: dataKeys[i].length + columnWidth });
      }
      workSheet["!cols"] = workSheetColumns;

      if (
        (calcName && calcCol && calcResult) ||
        (calcName && calcCol && calcFormula)
      ) {
        try {
          const calcNameSplit = calcCol.split(/(\d+)/);

          const getLastCol = (char: string) =>
            String.fromCharCode(char.charCodeAt(0) - 1);

          const calcNamePos = getLastCol(calcNameSplit[0]) + calcNameSplit[1];

          XLSX.utils.sheet_add_aoa(workSheet, [[calcName]], {
            origin: calcNamePos,
          });

          if (calcName && calcCol && calcResult) {
            console.log(calcResult);
            XLSX.utils.sheet_add_aoa(workSheet, [[calcResult]], {
              origin: calcCol,
            });
          } else if (calcName && calcCol && calcFormula) {
            XLSX.utils.sheet_set_array_formula(workSheet, calcCol, calcFormula);
          }
        } catch (error) {
          console.log("Error Calculation: ", error);
        }
      }

      //Create File
      const workBook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workBook, workSheet, sheetName || "Sheet1");
      XLSX.writeFile(workBook, `${sheetName}.xlsx`);
    } catch (error) {
      console.log("Error ExportSheet: ", error);
    }
  } else {
    let workSheet_tmp: XLSX.WorkSheet = XLSX.utils.json_to_sheet(
      exportData[0].data
    );
    let workSheet: XLSX.WorkSheet = XLSX.utils.sheet_to_json(workSheet_tmp, {
      header: 1,
    });

    exportData.slice(1).forEach((table) => {
      const { data, sheetName, calcName, calcCol, calcFormula, calcResult } =
        table;
      const dataKeys = Object.keys(data[0]);

      const tempSheetTable: XLSX.WorkSheet = XLSX.utils.json_to_sheet(data);

      const sheetTable: XLSX.WorkSheet = XLSX.utils.sheet_to_json(
        tempSheetTable,
        { header: 1 }
      );

      workSheet = workSheet.concat([""]).concat(sheetTable);
    });

    workSheet = XLSX.utils.aoa_to_sheet(workSheet, { skipHeader: true });

    const workBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workBook,
      workSheet,
      exportData[0].sheetName || "Sheet1"
    );
    
    XLSX.writeFile(workBook, `${exportData[0].sheetName}.xlsx`);
  }
};
