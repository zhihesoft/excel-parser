import xlsx from "node-xlsx";
import { importExcelItemMetaKey, importExcelItemMetaKeyAll } from "./excel.decorator";
import { excelDefaultParser } from "./excel.util";

/**
 * Represents the options for an Excel book.
 * keyRow: The row index of the key row. The key row contains the column names.
 */
export class ExcelBookOption {
    keyRow?: number;
}

/**
 * Represents an Excel book.
 */
export class ExcelBook {
    name: string = "";
    data: unknown[][] = [];
}

/**
 * Excel data
 */
export class ExcelFile {
    private books: ExcelBook[] = [];
    private keys: Map<string, number> = new Map();
    private cursor: number = 1;
    private rows: unknown[][] = [];

    public open(path: string) {
        this.books = xlsx.parse(path);
    }

    /**
     * 返回所有sheet的名称
     * @returns
     */
    public getSheetNames() {
        return this.books.map(book => book.name);
    }

    /**
     * 返回所有的key
     * @returns
     */
    public getSheetKeys() {
        const items = Array.from(this.keys.entries());
        items.sort((a, b) => a[1] - b[1]);
        return items.map(i => i[0]);
    }

    /**
     * Validates the data.
     * @returns {boolean} Returns true if the data is valid, false otherwise.
     */
    public validate() {
        if (!this.books) {
            return false;
        }
        if (this.books.length <= 0) {
            return false;
        }
        return true;
    }

    public getCursor() {
        return this.cursor;
    }

    /**
     * Sets the active book and updates the rows and cursor accordingly.
     * If the index is a number, it sets the active book based on the index.
     * If the index is a string, it finds the book with the matching name and sets it as the active book.
     * @param index - The index or name of the book to set as active.
     * @param option - Optional ExcelBookOption object.
     */
    public use(index: number | string, option?: ExcelBookOption) {
        const bookIndex = typeof index === "number" ? index : this.books.findIndex(book => book.name === index);
        this.rows = this.books[bookIndex].data;
        this.cursor = option?.keyRow ? option.keyRow : 0;
        const keyRow = this.rows[option?.keyRow ?? 0];
        this.keys.clear();
        keyRow.forEach((key, col) => {
            this.keys.set(String(key).trim(), col);
        });
    }

    /**
     * Retrieves the value of a cell in the Excel data.
     *
     * @param index - The index of the cell, either as a string or a number.
     * @param defaultValue - The default value to return if the cell is empty or undefined.
     * @returns The value of the cell as a string.
     */
    public getCell(index: string | number, defaultValue?: string): string | undefined {
        const col = this.parseIndex(index);
        if (col < 0) {
            return undefined;
        }
        const value = this.rows[this.cursor][col];
        if (!value) {
            return defaultValue;
        }
        // eslint-disable-next-line @typescript-eslint/no-base-to-string
        const ret = String(value).trim();
        return ret.trim();
    }

    /**
     * Parses the given instance by converting its properties based on metadata.
     * @param instance The instance to parse.
     */
    public parseTo(instance: unknown): unknown {
        for (const key in instance as any) {
            if (Object.prototype.hasOwnProperty.call(instance, key)) {
                const allMeta = Reflect.getMetadata(importExcelItemMetaKeyAll, instance as object, key);
                if (allMeta) {
                    const keys = this.getSheetKeys();
                    const obj: { [index: string]: string | undefined } = {};
                    for (const key of keys) {
                        obj[key] = this.getCell(key);
                    }
                    (instance as any)[key] = obj;
                    continue;
                }

                const meta = Reflect.getMetadata(importExcelItemMetaKey, instance as object, key);
                if (!meta) {
                    continue;
                }
                let convertor: (raw: string) => unknown = excelDefaultParser;
                if (meta.convertor) {
                    convertor = meta.convertor;
                }
                try {
                    const element = convertor.call(this, this.getCell(meta.column) ?? "");
                    (instance as any)[key] = element;
                } catch (err) {
                    // logger.error("parseTo failed", err, meta.column, this.getCell(meta.column));
                    throw err;
                }
            }
        }
        return instance;
    }

    /**
     * Moves the cursor to the next row in the Excel data.
     * @returns {boolean} True if there is a next row, false otherwise.
     */
    public next(skipBlank: boolean = true): boolean {
        if (this.cursor >= this.rows.length) {
            return false;
        }
        this.cursor++;
        if (skipBlank) {
            while (this.isBlankRow()) {
                if (this.cursor >= this.rows.length) {
                    return false;
                }
                this.cursor++;
            }
        }

        return this.cursor < this.rows.length;
    }

    /**
     * 当前行是否为空行（全是非显示字符）
     * @returns
     */
    public isBlankRow() {
        if (this.cursor >= this.rows.length) {
            return true;
        }
        for (const column of this.rows[this.cursor]) {
            const str = String(column).trim();
            if (str && str.length > 0) {
                return false;
            }
        }
        return true;
    }

    /**
     * Indicates whether the cursor has reached the end of the rows.
     * @returns {boolean} True if the cursor is at the end, false otherwise.
     */
    public get end(): boolean {
        return this.cursor >= this.rows.length;
    }

    private parseIndex(index: string | number): number {
        if (typeof index === "number") {
            if (index >= this.keys.size) {
                throw new Error(`Index ${index} out of range`);
            }
            return index;
        }
        if (!this.keys.has(index)) {
            // throw new Error(`Key ${index} not found`);
            return -1;
        }
        return this.keys.get(index) ?? -1;
    }
}
