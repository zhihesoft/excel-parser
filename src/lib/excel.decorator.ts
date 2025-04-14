import "reflect-metadata";
import { excelParser } from "./excel.util";

export const importExcelItemMetaKey = Symbol("import-excel-item-meta-key");
export const importExcelItemMetaKeyAll = Symbol("import-excel-item-meta-key-all");

export function ExcelColumn(column: string, convertor?: excelParser) {
    return Reflect.metadata(importExcelItemMetaKey, { column, convertor });
}

export function ExcelColumnAll() {
    return Reflect.metadata(importExcelItemMetaKeyAll, true);
}
