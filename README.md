# excel-parser

<p align="center">
  <a href="https://www.npmjs.com/package/@leadinvr/excel-parser">
    <img src="https://img.shields.io/npm/v/@leadinvr/excel-parser.svg?style=for-the-badge" alt="npm version" />
  </a>
  <a href="https://www.npmjs.com/package/@leadinvr/excel-parser">
    <img src="https://img.shields.io/npm/dt/@leadinvr/excel-parser.svg?style=for-the-badge" alt="npm total downloads" />
  </a>
  <a href="https://www.npmjs.com/package/@leadinvr/excel-parser">
    <img src="https://img.shields.io/npm/dm/@leadinvr/excel-parser.svg?style=for-the-badge" alt="npm monthly downloads" />
  </a>
  <a href="https://www.npmjs.com/package/@leadinvr/excel-parser">
    <img src="https://img.shields.io/npm/l/@leadinvr/excel-parser.svg?style=for-the-badge" alt="npm license" />
  </a>
</p>

# Features

-   Parse excel files

# Install

```bash
npm i @leadinvr/excel-parser
```

# Quick Start

### Parse an xlsx file

```js
try {
    const file = new ExcelFile();
    file.open(path);
    return file;
} catch (err) {
    logger.error(path, err);
    throw err;
}
```

### Select a book

```js
excel.use(0); // or excel.use("sheet name")
```

### Get all keys

```js
const keys = excel.getSheetKeys();
```

### Iterate rows

```js
while (excel.next()) {
    // get next row, and iterate all keys
    for (const key of keys) {
        const v = excel.getCell(key, "");
        if (isNullOrEmpty(v)) {
            continue;
        }
        if (!map.has(key)) {
            map.set(key, []);
        }
        map.get(key)!.push(v ?? "");
    }
}
```

### Decorators

* When call parseTo, parser use decorator to get excel cell from column

```js
/**
 * Example Data Class
 */
export class TestData {
    // name
    @ExcelColumn("name")
    name: string = "";

    @ExcelColumn("district")
    district: string = "";

    @ExcelColumn("type")
    type: string = "";

    @ExcelColumn("subtype")
    subtype: string = "";

    @ExcelColumnAll()
    properties: unknown = {};
}
```

* ExcelColumn

The field value will be cell under the argument key

* ExcelColumnAll

The field value will be a key-value object for all key in excel


### Parse row to an object

* It dependent on decorators

```js
while (excel.next()) {
    const obj = new TestData(); // Object class is any typescript class
    // the obj should has decorators on it's field
    excel.parseTo(obj);
}
```

