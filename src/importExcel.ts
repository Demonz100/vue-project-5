import * as XLSX from "xlsx";

export const importSheet = async (event: Event): Promise<object[]> => {
  const target = event.target as HTMLInputElement;
  let selectedFile = target.files!["0"];
  const importedData: object[] = [];

  const readFile = () => {
    return new Promise((resolve) => {
      //Convert selectedFile to/from binary
      try {
        let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFile);

        fileReader.onload = (event: Event) => {
          let data = fileReader.result;

          let workBook = XLSX.read(data, { type: "binary" });

          workBook.SheetNames.forEach((sheet) => {
            let rowObject: object[] = XLSX.utils.sheet_to_json(
              workBook.Sheets[sheet]
            );
            importedData.push(...rowObject);
          });
          resolve(importedData);
        };
      } catch (error) {
        console.log("Error Convert File: ", error);
      }
    });
  };

  const returnData = async () => {
    return await readFile().then(() => importedData);
  };

  return returnData();
};
