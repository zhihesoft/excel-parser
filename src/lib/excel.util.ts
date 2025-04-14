export type excelParser = (raw: string) => unknown;

/**
 * Parses the raw data from an Excel file and returns a string or number.
 *
 * @param raw The raw data to be parsed.
 * @returns The parsed data as a string or number.
 */
export function excelDefaultParser(raw: string): string | number | undefined {
    if (typeof raw === "number") {
        return Number(raw);
    }
    if (!raw) {
        return undefined;
    }
    return String(raw).trim();
}

/**
 * Parses the raw data from an Excel file and returns a string or number.
 *
 * @param raw The raw data to be parsed.
 * @returns The parsed data as a string or number.
 */
export function excelDefaultValue(defaultValue: string | number): excelParser {
    return raw => {
        if (!raw) {
            return defaultValue;
        }
        return excelDefaultParser(raw);
    };
}

/**
 * Parses a raw string into an array of values using a specified splitter.
 *
 * @param splitter The string used to split the raw input. Defaults to "[\\,，\r\n]+".
 * @returns An array of parsed values.
 */
export function excelArrayParserReg(splitter: string = "[\\,，\r\n]+"): excelParser {
    return raw => {
        if (!raw) {
            return [];
        }
        const ret = raw.split(new RegExp(splitter)).map(item => item.trim());
        return ret;
    };
}

/**
 * Parses a raw string into an array of values using a specified splitter.
 * @param splitter The delimiter used to split the raw string. Defaults to ",".
 * @returns An array of parsed values.
 */
export function excelArrayParser(splitter: string = ","): excelParser {
    return raw => {
        if (!raw) {
            return [];
        }
        const ret = raw.split(splitter).map(item => item.trim());
        return ret;
    };
}

/**
 * Parses a string representation of a date in Excel format and returns a Date object.
 * If the string contains a time component, it will be parsed as is. Otherwise, it will be treated as a numeric representation of a date.
 * @returns A Date object representing the parsed date.
 */
export function excelDateParser(raw: string): Date {
    const eternal = new Date(3000, 0);
    if (!raw) {
        return eternal;
    }
    if (raw == "长期" || !raw || raw.length <= 0) {
        return eternal;
    }
    if (typeof raw == "string" && (raw.includes(":") || raw.includes("-"))) {
        return new Date(raw);
    } else {
        return convertExcelDate(Number(raw));
    }
}

/**
 * Converts an Excel date to a JavaScript Date object.
 * @param excelDate - The Excel date to convert.
 * @returns The converted JavaScript Date object.
 */
export function convertExcelDate(excelDate: number): Date {
    const date = new Date(1899, 11, 30); // Excel 日期时间的起始日期为 1899 年 12 月 30 日
    const milliseconds = excelDate * 24 * 60 * 60 * 1000; // Excel 日期时间是自起始日期以来的天数，乘以一天的毫秒数
    date.setMilliseconds(date.getMilliseconds() + milliseconds);
    return date;
}
