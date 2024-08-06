# @connectk12/exceljs

This library provides a convenient way interact with Excel spreadsheet files. It contains of the following:
- Workbook functions to open and export Excel workbooks
- Data functions to help update cells, find rows of data, find data groups (multiple rows) for an entity, find the subsequent data group, etc.
- Validation functions to validate whether the Excel workbook matches an expected template with column headers, etc.
- Helper functions to handle typical scenarios that require matching data from various sources (sanitizeText, colLetterToNumber, etc.)

# Special thanks
The project code uses the NPM package [@zurmokeeper/exceljs](@zurmokeeper/exceljs) v4.4.7 and exports it. The @zurmokeeper/exceljs package is forked from [@exceljs/exceljs](https://github.com/exceljs/exceljs) v4.3.0.

Sincere thanks to all the developers of the @exceljs/exceljs and @zurmokeeper/exceljs project.

@connectk12/exceljs specifically uses @zurmokeeper/exceljs because of its capability in opening workbooks that are password-protected, as well as other bug fixes and features, since @exceljs/exceljs was no longer being supported by the team.

## Installation

To install the library, run the following command:

```bash
npm install @connectk12/exceljs
```

## Usage

This library contains the [@zurmokeeper/exceljs](@zurmokeeper/exceljs) library, so feel free to use the library as you would with the original library. The library also contains additional functions to help with common Excel tasks.

```javascript
import ExcelJS from '@connectk12/exceljs';
```


## Documentation

For detailed information on the available methods from @zurmokeeper/exceljs, please refer to the [@zurmokeeper/exceljs documentation](https://github.com/zurmokeeper/excelize) or [@exceljs/exceljs documentation](https://github.com/exceljs/exceljs).


## Get Started

Here are some examples to get you started:

### Example 1: Open Excel workbook, update cell, export workbook
```javascript
import ExcelJS from '@connectk12/exceljs';
import { updateCell } from '@connectk12/exceljs';
```

Notes: Regarding getting a worksheet from a workbook:
- Use `workbook.worksheets[n]` to get the nth worksheet in order, or `workbook.getWorksheet("Sheet 1")` to get the worksheet by name
- `workbook.getWorksheet(1)` returns the worksheet at index 1, which is not always the first worksheet

Inside an async function:
```javascript
const workbook = await openWorkbook({ workbookPathname: 'path/to/workbook.xlsx' });

// Get the first worksheet
const worksheet = workbook.worksheets[0]

// Get the first row and cell
const row = worksheet.getRow(1);
const cell = row.getCell(1);

// Update the cell with a value
updateCell(cell, 'Hello, World!');

// Export the workbook
const workbook = await exportWorkbook({
  workbook,
  outputPathname: 'path/to/exported/workbook.xlsx'
});
```

## Contributing

Contributions are welcome! If you find any issues or have suggestions for improvements, please open an issue or submit a pull request on the [GitHub repository](https://github.com/connectk12/exceljs).

## License

This library is licensed under the [MIT License](https://opensource.org/licenses/MIT).