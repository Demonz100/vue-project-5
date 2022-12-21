import * as XLSX from "xlsx-js-style";

export const exportSheet = (data: object[], sheetName?: string) => {
  const fileName = sheetName || 'Sheet1'
  const cellColor = {
    fill: { fgColor: { rgb: "1e81c6" } }
  };
  const headerColor = {
    font: { sz: 14, color: { rgb: "ffffff" }, bold: true }
  };
  const columnWidth = 20;
  const cellBorder = {
    border: {
      top: { style: "thin", color: { rgb: "00000" } },
      right: { style: "thin", color: { rgb: "00000" } },
      bottom: { style: "thin", color: { rgb: "00000" } },
      left: { style: "thin", color: { rgb: "00000" } }
    }
  };

  try {
    const dataKeys = Object.keys(data[0]);

    const workBook = XLSX.utils.book_new();
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
          ...cellBorder
        };
        if (col === 0) {
          //Format first column
          workSheet[cellRef].s = {
            alignment: { horizontal: "center" },
            ...cellBorder
          };
        }
        if (row === 0) {
          // Format first row
          workSheet[cellRef].s = {
            ...workSheet[cellRef].s,
            ...cellColor,
            ...cellBorder,
            ...headerColor
          };
        }
      }
    }

    const workSheetColumns = [];
    for (let i = 0; i < dataKeys.length; i++) {
      workSheetColumns.push({ wch: dataKeys[i].length + columnWidth });
    }
    workSheet["!cols"] = workSheetColumns;

    //Create File
    XLSX.utils.book_append_sheet(workBook, workSheet, fileName);
    XLSX.writeFile(workBook, `${fileName}.xlsx`);
  } catch (error) {
    console.log("Error ExportSheet: ", error)
  }
};
