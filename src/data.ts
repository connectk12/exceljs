import ExcelJS from "@zurmokeeper/exceljs";
import { sanitizeText } from "./helpers";

/* List of functions in this file:
 * updateCell
 * findMatchingRowById
 * findMatchingRowByName
 * findNextRecordSetRowById
 * findNextRecordSetRowByName
 * getRowsOfMatchingRecordSet
 * createNewRowAfterRecordSet
 */

/**
 * Updates the cell with the provided value and options.
 * It can overwrite the cell value, update the cell style, fill color, number format, and note.
 *
 * @param cell - The cell to be updated.
 * @param value - The value to be set in the cell.
 * @param opts - Options to customize the cell update process.
 * @param opts.overwrite - Whether to overwrite the cell value if it already exists. Default is true.
 * @param opts.style - The style to be applied to the cell.
 * @param opts.fillColor - The fill color to be applied to the cell.
 * @param opts.removeFillColor - Whether to remove the fill color from the cell.
 * @param opts.numFmt - The number format to be applied to the cell.
 * @param opts.note - The note to be added to the cell.
 *
 * @returns {void}
 *
 * @example
 * updateCell(cell, "Hello, World!");
 */
export const updateCell = (
  cell: ExcelJS.Cell,
  value?: ExcelJS.CellValue,
  opts?: {
    overwrite?: boolean;
    style?: Partial<ExcelJS.Style>;
    fillColor?: string;
    removeFillColor?: boolean;
    numFmt?: string;
    note?: string;
  }
): void => {
  const overwrite = opts?.overwrite ?? true;
  if (!overwrite && cell.value && cell.value !== "") return;

  cell.value = value;
  if (opts?.style) {
    cell.style = {
      ...cell.style,
      ...opts.style,
    };
  }
  if (opts?.fillColor) {
    cell.style = {
      ...cell.style,
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: opts.fillColor },
      },
    };
  }
  if (opts?.removeFillColor) {
    cell.style = {
      ...cell.style,
      fill: {
        type: "pattern",
        pattern: "none",
      },
    };
  }
  if (opts?.numFmt) {
    cell.numFmt = opts.numFmt;
  }
  if (opts?.note) {
    cell.note = {
      texts: [{ text: opts.note }],
      editAs: "oneCells",
    };
  }
};

/**
 * Finds the matching row in the worksheet based on the provided ID and options.
 * It can find the exact match or a similar match based on the last digits of the ID.
 *
 * @param worksheet - The worksheet to search for the matching row.
 * @param startRowNumber - The starting row number to search for the matching row.
 * @param lastRowNumber - The last row number to search for the matching row.
 * @param id - The ID to search for in the worksheet.
 * @param lookupCol - The column to search for the ID in the worksheet.
 * @param lookupCondition - The condition to match the lookup value.
 * @param opts - Options to customize the search process.
 * @param opts.findSimilarMatchWithLastDigits - Whether to find a similar match based on the last digits of the ID.
 *
 * @returns {ExcelJS.Row | undefined} - The matching row if found, otherwise undefined.
 *
 * @example
 * const row = findMatchingRowById({
 *   worksheet,
 *   startRowNumber: 2,
 *   lastRowNumber: 10,
 *   id: "12345",
 *   lookupCol: "A",
 *   opts: { findSimilarMatchWithLastDigits: true }
 * });
 */
export const findMatchingRowById = ({
  worksheet,
  startRowNumber,
  lastRowNumber,
  id,
  lookupCol,
  lookupCondition,
  opts,
}: {
  worksheet: ExcelJS.Worksheet;
  startRowNumber: number;
  lastRowNumber: number;
  id: string;
  lookupCol: string;
  lookupCondition?: (currentRowValue: string) => boolean;
  opts?: {
    findSimilarMatchWithLastDigits?: boolean;
  };
}): ExcelJS.Row | undefined => {
  const rows = worksheet.getRows(startRowNumber, lastRowNumber);
  if (!rows) {
    console.log("No matching rows found");
    return undefined;
  }

  const row = rows.find((row) => {
    const _id = sanitizeText(id);
    const _lookupValue = sanitizeText(row.getCell(lookupCol).value?.toString());
    return _id === _lookupValue;
  });

  if (!row && lookupCondition) {
    return rows.find((row) => {
      const currentRowId = sanitizeText(
        row.getCell(lookupCol).value?.toString()
      );
      return currentRowId && lookupCondition(currentRowId);
    });
  } else if (!row && opts?.findSimilarMatchWithLastDigits) {
    return rows.find((row) => {
      const currentRowId = sanitizeText(
        row.getCell(lookupCol).value?.toString()
      );
      return currentRowId?.endsWith(id);
    });
  }
  return row;
};

/**
 * Finds the matching row in the worksheet based on the provided name and options.
 * It can find the exact match or a similar match based on the name (although an exact match is still performed first).
 * The name can be split into first name and last name or combined with a delimiter.
 *
 * @param worksheet - The worksheet to search for the matching row.
 * @param startRowNumber - The starting row number to search for the matching row.
 * @param lastRowNumber - The last row number to search for the matching row.
 * @param nameValues - The name values to search for in the worksheet.
 * @param lookupCols - The columns to search for the name in the worksheet.
 * @param opts - Options to customize the search process.
 * @param opts.findSimilarMatch - Whether to find a similar match based on the name if the exact match is not found. Default is true.
 *
 * @returns {ExcelJS.Row | undefined} - The matching row if found, otherwise undefined.
 *
 * @example
 * const row = findMatchingRowByName({
 *   worksheet,
 *   startRowNumber: 2,
 *   lastRowNumber: 10,
 *   nameValues: [{ firstName: "John", lastName: "Doe" }],
 *   lookupCols: { firstName: "A", lastName: "B" },
 *   opts: { findSimilarMatch: true }
 * });
 *
 */

export const findMatchingRowByName = ({
  worksheet,
  startRowNumber,
  lastRowNumber,
  nameValues,
  lookupCols,
  opts,
}: {
  worksheet: ExcelJS.Worksheet;
  startRowNumber: number;
  lastRowNumber: number;
  nameValues: {
    firstName?: string;
    lastName?: string;
  }[];
  lookupCols:
    | {
        firstName: string;
        lastName: string;
      }
    | {
        name: string;
        order: "FIRST_NAME LAST_NAME" | "LAST_NAME FIRST_NAME";
        delimiter?: string;
      };
  opts?: {
    findSimilarMatch?: boolean;
  };
}): ExcelJS.Row | undefined => {
  const _findSimilarMatch = opts?.findSimilarMatch ?? true;
  const sanitizedNameValues = nameValues.map((nameValue) => ({
    lastName: sanitizeText(nameValue.lastName, { uppercase: true }),
    firstName: sanitizeText(nameValue.firstName, { uppercase: true }),
  }));
  const rows = worksheet.getRows(startRowNumber, lastRowNumber);
  if (!rows) {
    // throw new Error("No matching rows found");
    console.log("No matching rows found");
    return undefined;
  }

  // Find exact match first
  let row = rows.find((row) => {
    let worksheetLastName = undefined;
    let worksheetFirstName = undefined;

    if ("lastName" in lookupCols && "firstName" in lookupCols) {
      worksheetLastName = sanitizeText(
        row.getCell(lookupCols.lastName).value?.toString(),
        { uppercase: true }
      );
      worksheetFirstName = sanitizeText(
        row.getCell(lookupCols.firstName).value?.toString(),
        { uppercase: true }
      );
    } else if ("name" in lookupCols) {
      const currentRowName = sanitizeText(
        row.getCell(lookupCols.name).value?.toString(),
        { uppercase: true }
      );
      if (currentRowName) {
        const nameArray = currentRowName.split(lookupCols.delimiter ?? " ");
        if (
          nameArray.length > 1 &&
          lookupCols.order === "FIRST_NAME LAST_NAME"
        ) {
          worksheetFirstName = sanitizeText(nameArray[0], { uppercase: true });
          worksheetLastName = sanitizeText(nameArray[1], { uppercase: true });
        } else if (
          nameArray.length > 1 &&
          lookupCols.order === "LAST_NAME FIRST_NAME"
        ) {
          worksheetFirstName = sanitizeText(nameArray[1], { uppercase: true });
          worksheetLastName = sanitizeText(nameArray[0], { uppercase: true });
        }
      }
    }

    if (!worksheetLastName || !worksheetFirstName) {
      return false;
    }

    const match = sanitizedNameValues.some(
      (nameValue) =>
        worksheetLastName === nameValue.lastName &&
        worksheetFirstName === nameValue.firstName
    );
    return match;
  });

  // Find similar match if exact match not found
  if (!row && _findSimilarMatch) {
    row = rows.find((row) => {
      let worksheetLastName = undefined;
      let worksheetFirstName = undefined;
      if ("lastName" in lookupCols && "firstName" in lookupCols) {
        worksheetLastName = sanitizeText(
          row.getCell(lookupCols.lastName).value?.toString(),
          { uppercase: true }
        );
        worksheetFirstName = sanitizeText(
          row.getCell(lookupCols.firstName).value?.toString(),
          { uppercase: true }
        );
      } else if ("name" in lookupCols) {
        const currentRowName = sanitizeText(
          row.getCell(lookupCols.name).value?.toString(),
          { uppercase: true }
        );
        if (currentRowName) {
          const nameArray = currentRowName.split(lookupCols.delimiter ?? " ");
          if (
            nameArray.length > 1 &&
            lookupCols.order === "FIRST_NAME LAST_NAME"
          ) {
            worksheetFirstName = sanitizeText(nameArray[0], {
              uppercase: true,
            });
            worksheetLastName = sanitizeText(nameArray[1], { uppercase: true });
          } else if (
            nameArray.length > 1 &&
            lookupCols.order === "LAST_NAME FIRST_NAME"
          ) {
            worksheetFirstName = sanitizeText(nameArray[1], {
              uppercase: true,
            });
            worksheetLastName = sanitizeText(nameArray[0], { uppercase: true });
          }
        }
      }
      if (!worksheetLastName || !worksheetFirstName) {
        return false;
      }

      const match = sanitizedNameValues.some(
        (nameValue) =>
          nameValue.lastName &&
          nameValue.lastName !== "" &&
          nameValue.firstName &&
          nameValue.firstName !== "" &&
          ((worksheetLastName.includes(nameValue.lastName) &&
            worksheetFirstName.includes(nameValue.firstName)) ||
            (nameValue.lastName?.includes(worksheetLastName) &&
              nameValue.firstName?.includes(worksheetFirstName)))
      );
      return match;
    });
  }

  return row;
};

/**
 * Finds the last row in the record set based on the starting row and the lookup value.
 * It searches for the next row that does not match the lookup value.
 *
 * @param worksheet - The worksheet to search for the last row in the record set.
 * @param startRowNumber - The starting row to search for the last row in the record set.
 * @param lastRowNumber - The last row number to search for the last row in the record set.
 * @param lookupCol - The column to search for the lookup value in the worksheet.
 * @param lookupValue - The value to match the lookup value.
 * @param lookupCondition - The condition to match the lookup value.
 *
 * @returns {ExcelJS.Row | undefined} - The last row in the record set if found, otherwise undefined.
 */
export const findLastRowInRecordSet = ({
  worksheet,
  startRowNumber,
  lastRowNumber,
  lookupCol,
  lookupValue,
  lookupCondition,
}: {
  worksheet: ExcelJS.Worksheet;
  startRowNumber: number;
  lastRowNumber: number;
  lookupCol: string;
  lookupValue: ExcelJS.CellValue;
  lookupCondition?: (currentRowValue: string) => boolean;
}): ExcelJS.Row | undefined => {
  const rows = worksheet.getRows(
    startRowNumber,
    lastRowNumber - startRowNumber + 1
  );
  if (!rows) {
    console.log("No matching rows found");
    return undefined;
  }

  let nextMatchingRow = rows.find((row) => {
    let currentRowValue = sanitizeText(
      row.getCell(lookupCol).value?.toString()
    );
    if (lookupCondition) {
      return (
        !currentRowValue ||
        (currentRowValue && !lookupCondition(currentRowValue))
      );
    }
    return currentRowValue !== lookupValue;
  });
  const lastRowInRecordSet = nextMatchingRow
    ? worksheet.getRow(nextMatchingRow.number - 1)
    : worksheet.getRow(lastRowNumber);
  return lastRowInRecordSet;
};

/**
 * Finds the next record set in the worksheet based on the starting row by id.
 * It searches for the next row that does not match id.
 *
 * @param worksheet - The worksheet to search for the next record set.
 * @param startRowNumber - The starting row number to search for the next record set.
 * @param lastRowNumber - The last row number to search for the next record set.
 * @param id - The id to match the id value.
 * @param lookupCol - The column to search for the id in the worksheet.
 * @param opts - Options to customize the search process.
 *
 * @returns {ExcelJS.Row | undefined} - The next record set if found, otherwise undefined.
 */
export const findNextRecordSetRowById = ({
  worksheet,
  startRowNumber,
  lastRowNumber,
  id,
  lookupCol,
  opts,
}: {
  worksheet: ExcelJS.Worksheet;
  startRowNumber: number;
  lastRowNumber: number;
  id: string;
  lookupCol: string;
  opts?: {
    findSimilarMatchWithLastDigits?: boolean;
  };
}): ExcelJS.Row | undefined => {
  const rows = worksheet.getRows(startRowNumber, lastRowNumber);
  if (!rows) {
    // throw new Error("No matching rows found");
    console.log("No matching rows found");
    return undefined;
  }

  let nextMatchingRow = rows.find((row) => {
    let currentRowId = sanitizeText(row.getCell(lookupCol).value?.toString());
    if (!currentRowId) {
      return true;
    }
    return (
      currentRowId === id &&
      (opts?.findSimilarMatchWithLastDigits ? currentRowId.endsWith(id) : true)
    );
  });

  return nextMatchingRow;
};

/**
 * Finds the next record set in the worksheet based on the starting row by name.
 * It searches for the next row that does not match the name.
 * The name can be split into first name and last name or combined with a delimiter.
 * It can also include conditions to match additional columns.
 *
 * @param worksheet - The worksheet to search for the next record set.
 * @param startRowNumber - The starting row to search for the next record set.
 * @param lastRowNumber - The last row number to search for the next record set.
 * @param lookupCols - The columns to search for the name in the worksheet.
 * @param conditions - The conditions to match additional columns.
 *
 * @returns {ExcelJS.Row | undefined} - The next record set if found, otherwise undefined.
 */
export const findNextRecordSetRowByName = ({
  worksheet,
  startRowNumber,
  lastRowNumber,
  lookupCols,
  conditions,
}: {
  worksheet: ExcelJS.Worksheet;
  startRowNumber: number;
  lastRowNumber: number;
  lookupCols:
    | {
        firstName: string;
        lastName: string;
      }
    | {
        name: string;
        order: "FIRST_NAME LAST_NAME" | "LAST_NAME FIRST_NAME";
        delimiter?: string;
      };
  conditions?: {
    col: string;
    condition: (cell: ExcelJS.Cell) => boolean;
  }[];
}): ExcelJS.Row | undefined => {
  const startRow = worksheet.getRow(startRowNumber);
  let currentLastName = undefined;
  let currentFirstName = undefined;
  if ("lastName" in lookupCols && "firstName" in lookupCols) {
    currentLastName = sanitizeText(
      startRow.getCell(lookupCols.lastName).value?.toString()
    );
    currentFirstName = sanitizeText(
      startRow.getCell(lookupCols.firstName).value?.toString()
    );
  } else if ("name" in lookupCols) {
    const currentEmployeeName = sanitizeText(
      startRow.getCell(lookupCols.name).value?.toString(),
      { removeSpecialChars: false, removeWhitespace: false }
    );

    if (currentEmployeeName) {
      const nameArray = currentEmployeeName.split(lookupCols.delimiter ?? " ");
      if (nameArray.length > 1 && lookupCols.order === "FIRST_NAME LAST_NAME") {
        currentFirstName = nameArray[0]?.trim();
        currentLastName = nameArray[1]?.trim();
      } else if (
        nameArray.length > 1 &&
        lookupCols.order === "LAST_NAME FIRST_NAME"
      ) {
        currentFirstName = nameArray[1]?.trim();
        currentLastName = nameArray[0]?.trim();
      }
    }
  }

  const rows = worksheet.getRows(startRow.number, lastRowNumber);
  if (!rows) {
    // throw new Error("No matching rows found");
    console.log("No matching rows found");
    return undefined;
  }

  let nextMatchingRow = rows.find((row) => {
    let nextRowLastName = undefined;
    let nextRowFirstName = undefined;
    if ("lastName" in lookupCols && "firstName" in lookupCols) {
      nextRowLastName = sanitizeText(
        row.getCell(lookupCols.lastName).value?.toString()
      );
      nextRowFirstName = sanitizeText(
        row.getCell(lookupCols.firstName).value?.toString()
      );
    } else if ("name" in lookupCols) {
      const currentRowName = sanitizeText(
        row.getCell(lookupCols.name).value?.toString(),
        { removeSpecialChars: false, removeWhitespace: false }
      );
      if (currentRowName) {
        const nameArray = currentRowName.split(lookupCols.delimiter ?? " ");
        if (
          nameArray.length > 1 &&
          lookupCols.order === "FIRST_NAME LAST_NAME"
        ) {
          nextRowFirstName = nameArray[0]?.trim();
          nextRowLastName = nameArray[1]?.trim();
        } else if (
          nameArray.length > 1 &&
          lookupCols.order === "LAST_NAME FIRST_NAME"
        ) {
          nextRowFirstName = nameArray[1]?.trim();
          nextRowLastName = nameArray[0]?.trim();
        }
      }
    }
    if (!nextRowLastName || !nextRowFirstName) {
      return true;
    }
    return (
      (nextRowLastName !== currentLastName ||
        nextRowFirstName !== currentFirstName) &&
      (conditions
        ? conditions?.every((condition) => {
            const cell = row.getCell(condition.col);
            return condition.condition(cell);
          })
        : true)
    );
  });

  return nextMatchingRow;
};

/**
 * Retrieves a record set (contiguous group of rows) from the worksheet that match a cell value at the lookup column and optional conditions
 *
 * @param worksheet - The worksheet to search for the matching rows.
 * @param startRowNumber - The starting row number to search for the matching rows.
 * @param lookupCol - The column to search for the lookup value in the worksheet.
 * @param lookupValue - The value to match the lookup value.
 * @param lookupCondition - The condition to match the lookup value.
 * @param opts - Options to customize the search process.
 * @param opts.findSimilarMatchWithLastDigits - Whether to find a similar match based on the last digits of the lookup value.
 *
 * @returns {ExcelJS.Row[]} - Returns the rows of the matching record set
 *
 * @example
 * const rows = getRowsOfMatchingRecordSet({
 *   worksheet,
 *   startRowNumber: 2,
 *   lookupCol: "A",
 *   lookupValue: "12345",
 *   opts: { findSimilarMatchWithLastDigits: true }
 * });
 */

export const getRowsOfMatchingRecordSet = ({
  worksheet,
  startRowNumber,
  lastRowNumber,
  lookupCol,
  lookupValue,
  lookupCondition,
  opts,
}: {
  worksheet: ExcelJS.Worksheet;
  startRowNumber: number;
  lastRowNumber: number;
  lookupCol: string;
  lookupValue: string;
  lookupCondition?: (currentRowValue: string) => boolean;
  opts?: {
    findSimilarMatchWithLastDigits?: boolean;
  };
}): ExcelJS.Row[] | undefined => {
  let rows: ExcelJS.Row[] | undefined = undefined;
  const startRow = findMatchingRowById({
    worksheet,
    startRowNumber,
    lastRowNumber,
    lookupCol,
    lookupCondition,
    id: lookupValue,
    opts,
  });
  if (startRow) {
    const lastRow = findLastRowInRecordSet({
      worksheet,
      startRowNumber: startRow.number,
      lastRowNumber,
      lookupCol,
      lookupValue,
      lookupCondition:
        lookupCondition ??
        (opts?.findSimilarMatchWithLastDigits
          ? (currentRowValue) => currentRowValue?.endsWith(lookupValue)
          : undefined),
    });
    if (startRow && lastRow) {
      rows = worksheet.getRows(
        startRow.number,
        lastRow.number - startRow.number + 1
      );
    }
  }
  return rows;
};

/**
 * Retrieves a set of rows from the worksheet that match a cell value at the lookup column and optional conditions
 * and returns the newly added row
 * @param worksheet - The worksheet to search for the matching rows.
 * @param startRowNumber - The starting row number to search for the matching rows.
 * @param lookupCol - The column to search for the lookup value in the worksheet.
 *
 * @returns {ExcelJS.Row} - Returns the last row of the record set
 *
 * @example
 * const newRow = createNewRowAfterRecordSet({
 *   worksheet,
 *   startRowNumber: 2,
 *   lookupCol: "A"
 * });
 */
export const createNewRowAfterRecordSet = ({
  worksheet,
  startRowNumber,
  lookupCol,
}: {
  worksheet: ExcelJS.Worksheet;
  startRowNumber: number;
  lookupCol: string;
}): ExcelJS.Row => {
  // Get rows for employee data
  let prepRowNumber = startRowNumber;
  let prepCellFieldValue = worksheet
    .getRow(prepRowNumber)
    .getCell(lookupCol).value;
  while (prepCellFieldValue !== null && prepCellFieldValue !== "") {
    prepRowNumber++;
    prepCellFieldValue = worksheet
      .getRow(prepRowNumber)
      .getCell(lookupCol).value;
  }
  const previousRecordSet = prepRowNumber - 1;

  // Duplicate previous row
  worksheet.insertRow(previousRecordSet + 1, [], "i+");
  const newRow = worksheet.getRow(previousRecordSet + 1);
  newRow.eachCell({ includeEmpty: true }, (cell) => {
    cell.value = "";
  });
  return newRow;
};
