import { Workbook } from "exceljs";

export interface IColumn<T> {
  columnName: string;
  key: string;
  valueMapper?: (value: T) => string;
}

export interface ISheetData<T> {
  sheetName: string;
  table: IColumn<any>[];
  data?: T[];
}

interface ISheet<T> {
  sheetName: string;
  table: IColumn<any>[];
  data: T[];
}

interface IExcelInstance {
  addSheet: <T>(sheetData: ISheetData<T>) => IExcelInstance;
  setData: <T>(targetSheet: string, data: T[]) => IExcelInstance;
  insertData: <T>(targetSheet: string, data: T | T[]) => IExcelInstance;
  download: (fileName?: string) => Promise<void>;
}

function defaultColumnValueMapper(value: any): string {
  if (!value) return "";

  return value.toString();
}

export function excel(
  fileName?: string,
  defaultWorkBook?: ISheetData<any>[]
): IExcelInstance {
  const defaultFileName = fileName;
  const sheets: ISheet<any>[] =
    defaultWorkBook?.map(({ sheetName, table, data }) => ({
      sheetName,
      table,
      data: data || [],
    })) || [];
  const sheetDic: Record<string, ISheet<any>> = sheets.reduce(
    (sheetDic, sheet) => {
      sheetDic[sheet.sheetName] = sheet;
      return sheetDic;
    },
    {} as Record<string, ISheet<any>>
  );

  function getSheet<T>(sheetName: string): ISheet<T> {
    const targetSheet = sheetDic[sheetName];

    if (!targetSheet) {
      throw new Error(
        `정의된 시트를 찾을 수 없습니다. 시트 이름 : ${sheetName}`
      );
    }

    return targetSheet;
  }

  function addSheet<T>(
    this: IExcelInstance,
    { sheetName, table, data }: ISheetData<T>
  ): IExcelInstance {
    const sheet = {
      sheetName,
      table,
      data: data || [],
    };

    sheets.push(sheet);
    sheetDic[sheetName] = sheet;

    return this;
  }

  function setData<T>(
    this: IExcelInstance,
    sheetName: string,
    data: T[]
  ): IExcelInstance {
    try {
      const targetSheet = getSheet(sheetName);
      targetSheet.data = data;
    } catch (e) {
      console.error(e);
    }

    return this;
  }

  function insertData<T>(
    this: IExcelInstance,
    sheetName: string,
    data: T | T[]
  ): IExcelInstance {
    try {
      const targetSheet = getSheet(sheetName);
      if (Array.isArray(data)) {
        targetSheet.data = [...targetSheet.data, ...data];
      } else {
        targetSheet.data = [...targetSheet.data, data];
      }
    } catch (e) {
      console.error(e);
    }

    return this;
  }

  async function download(fileName?: string): Promise<void> {
    const workbook = new Workbook();
    sheets.forEach(({ sheetName, table, data }) => {
      const sheet = workbook.addWorksheet(sheetName);

      table.forEach(({ columnName, key, valueMapper }, index) => {
        sheet.getColumn(index + 1).values = [
          columnName,
          ...data
            .map((data) => data[key])
            .map(valueMapper || defaultColumnValueMapper),
        ];
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();

    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    const downloadFileName =
      fileName ||
      defaultFileName ||
      `${crypto.randomUUID().split("-")[0]}.xlsx`;
    anchor.download = downloadFileName.includes(".")
      ? downloadFileName
      : `${downloadFileName}.xlsx`;
    anchor.href = url;
    anchor.click();

    window.URL.revokeObjectURL(url);
  }

  return {
    addSheet,
    setData,
    insertData,
    download,
  };
}
