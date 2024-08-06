import ExcelJS from "@zurmokeeper/exceljs";
import { sanitizeText } from "./helpers";

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
) => {
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

// Find first and last rows for employees
export const findFirstRowOfDataGroup = ({
  worksheet,
  startRowNumber,
  identifierCol,
  identifierCondition,
}: {
  worksheet: ExcelJS.Worksheet;
  startRowNumber: number;
  identifierCol: string;
  identifierCondition?: (identifierValue: string) => boolean;
}) => {
  // Get rows for employee data
  let rowNumber = startRowNumber;
  let identifierValue = worksheet
    .getRow(rowNumber)
    .getCell(identifierCol)
    .value?.toString();
  while (
    identifierCondition
      ? identifierValue
        ? identifierCondition(identifierValue)
        : false
      : identifierValue && identifierValue !== ""
  ) {
    rowNumber--;
    identifierValue = worksheet
      .getRow(rowNumber)
      .getCell(identifierCol)
      .value?.toString();
  }
  const firstRowNumber = rowNumber + 1;
  return firstRowNumber;
};

export const findLastRowOfDataGroup = ({
  worksheet,
  startRowNumber,
  identifierCol,
  identifierCondition,
}: {
  worksheet: ExcelJS.Worksheet;
  startRowNumber: number;
  identifierCol: string;
  identifierCondition?: (identifierValue: string) => boolean;
}) => {
  // Get rows for employee data
  let rowNumber = startRowNumber;
  let identifierValue = worksheet
    .getRow(rowNumber)
    .getCell(identifierCol)
    .value?.toString();
  while (
    identifierCondition
      ? identifierValue
        ? identifierCondition(identifierValue)
        : false
      : identifierValue && identifierValue !== ""
  ) {
    rowNumber++;
    identifierValue = worksheet
      .getRow(rowNumber)
      .getCell(identifierCol)
      .value?.toString();
  }
  const lastRowNumber = rowNumber - 1;
  return lastRowNumber;
};

// Create new row for employees not found
export const createNewRowAfterDataGroup = ({
  worksheet,
  startRowNumber,
  identifierCol,
}: {
  worksheet: ExcelJS.Worksheet;
  startRowNumber: number;
  identifierCol: string;
}) => {
  // Get rows for employee data
  let unknownEmployeePrepRowNumber = startRowNumber;
  let unknownEmployeePopulatedFieldValue = worksheet
    .getRow(unknownEmployeePrepRowNumber)
    .getCell(identifierCol).value;
  while (
    unknownEmployeePopulatedFieldValue !== null &&
    unknownEmployeePopulatedFieldValue !== ""
  ) {
    unknownEmployeePrepRowNumber++;
    unknownEmployeePopulatedFieldValue = worksheet
      .getRow(unknownEmployeePrepRowNumber)
      .getCell(identifierCol).value;
  }
  const unknownEmployeePrepLastEmployeeRow = unknownEmployeePrepRowNumber - 1;

  // Duplicate previous row
  worksheet.insertRow(unknownEmployeePrepLastEmployeeRow + 1, [], "i+");
  const newRow = worksheet.getRow(unknownEmployeePrepLastEmployeeRow + 1);
  newRow.eachCell({ includeEmpty: true }, (cell) => {
    cell.value = "";
  });
  return newRow;
};

// Find matching row by ID
export const findMatchingRowById = ({
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
}) => {
  const employeeRows = worksheet.getRows(startRowNumber, lastRowNumber);
  if (!employeeRows) {
    console.log("No employee rows found");
    return undefined;
  }

  const row = employeeRows.find(
    (row) => id === sanitizeText(row.getCell(lookupCol).value?.toString())
  );

  if (!row && opts?.findSimilarMatchWithLastDigits) {
    return employeeRows.find((row) => {
      const worksheetEmployeeId = sanitizeText(
        row.getCell(lookupCol).value?.toString()
      );
      return worksheetEmployeeId?.endsWith(id);
    });
  }
  return row;
};

// Find matching row by name
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
}) => {
  const _findSimilarMatch = opts?.findSimilarMatch ?? true;
  const sanitizedNameValues = nameValues.map((nameValue) => ({
    lastName: sanitizeText(nameValue.lastName),
    firstName: sanitizeText(nameValue.firstName),
  }));
  const employeeRows = worksheet.getRows(startRowNumber, lastRowNumber);
  if (!employeeRows) {
    // throw new Error("No employee rows found");
    console.log("No employee rows found");
    return undefined;
  }

  // Find exact match first
  let row = employeeRows.find((row) => {
    let worksheetEmployeeLastName = undefined;
    let worksheetEmployeeFirstName = undefined;

    if ("lastName" in lookupCols && "firstName" in lookupCols) {
      worksheetEmployeeLastName = sanitizeText(
        row.getCell(lookupCols.lastName).value?.toString()
      );
      worksheetEmployeeFirstName = sanitizeText(
        row.getCell(lookupCols.firstName).value?.toString()
      );
    } else if ("name" in lookupCols) {
      const worksheetEmployeeName = sanitizeText(
        row.getCell(lookupCols.name).value?.toString()
      );
      if (worksheetEmployeeName) {
        const nameArray = worksheetEmployeeName.split(
          lookupCols.delimiter ?? " "
        );
        if (
          nameArray.length > 1 &&
          lookupCols.order === "FIRST_NAME LAST_NAME"
        ) {
          worksheetEmployeeFirstName = nameArray[0]?.trim();
          worksheetEmployeeLastName = nameArray[1]?.trim();
        } else if (
          nameArray.length > 1 &&
          lookupCols.order === "LAST_NAME FIRST_NAME"
        ) {
          worksheetEmployeeFirstName = nameArray[1]?.trim();
          worksheetEmployeeLastName = nameArray[0]?.trim();
        }
      }
    }

    if (!worksheetEmployeeLastName || !worksheetEmployeeFirstName) {
      return false;
    }

    return sanitizedNameValues.some(
      (nameValue) =>
        worksheetEmployeeLastName === nameValue.lastName &&
        worksheetEmployeeFirstName === nameValue.firstName
    );
  });

  // Find similar match if exact match not found
  if (!row && _findSimilarMatch) {
    row = employeeRows.find((row) => {
      let worksheetEmployeeLastName = undefined;
      let worksheetEmployeeFirstName = undefined;
      if ("lastName" in lookupCols && "firstName" in lookupCols) {
        worksheetEmployeeLastName = sanitizeText(
          row.getCell(lookupCols.lastName).value?.toString()
        );
        worksheetEmployeeFirstName = sanitizeText(
          row.getCell(lookupCols.firstName).value?.toString()
        );
      } else if ("name" in lookupCols) {
        const worksheetEmployeeName = sanitizeText(
          row.getCell(lookupCols.name).value?.toString()
        );
        if (worksheetEmployeeName) {
          const nameArray = worksheetEmployeeName.split(
            lookupCols.delimiter ?? " "
          );
          if (
            nameArray.length > 1 &&
            lookupCols.order === "FIRST_NAME LAST_NAME"
          ) {
            worksheetEmployeeFirstName = nameArray[0]?.trim();
            worksheetEmployeeLastName = nameArray[1]?.trim();
          } else if (
            nameArray.length > 1 &&
            lookupCols.order === "LAST_NAME FIRST_NAME"
          ) {
            worksheetEmployeeFirstName = nameArray[1]?.trim();
            worksheetEmployeeLastName = nameArray[0]?.trim();
          }
        }
      }
      if (!worksheetEmployeeLastName || !worksheetEmployeeFirstName) {
        return false;
      }

      return sanitizedNameValues.some(
        (nameValue) =>
          nameValue.lastName &&
          nameValue.lastName !== "" &&
          nameValue.firstName &&
          nameValue.firstName !== "" &&
          ((worksheetEmployeeLastName.includes(nameValue.lastName) &&
            worksheetEmployeeFirstName.includes(nameValue.firstName)) ||
            (nameValue.lastName?.includes(worksheetEmployeeLastName) &&
              nameValue.firstName?.includes(worksheetEmployeeFirstName)))
      );
    });
  }

  return row;
};

// Find next employee row that does not match current employee by id
export const findNextDifferentEmployeeRowById = ({
  worksheet,
  currentEmployeeRow,
  lastRowNumber,
  lookupCol,
  conditions,
  opts,
}: {
  worksheet: ExcelJS.Worksheet;
  currentEmployeeRow: ExcelJS.Row;
  lastRowNumber: number;
  lookupCol: string;
  conditions?: {
    col: string;
    condition: (cell: ExcelJS.Cell) => boolean;
  }[];
  opts?: {
    findSimilarMatchWithLastDigits?: boolean;
  };
}) => {
  const currentEmployeeId = sanitizeText(
    currentEmployeeRow.getCell(lookupCol).value?.toString()
  );
  const employeeRows = worksheet.getRows(
    currentEmployeeRow.number,
    lastRowNumber
  );
  if (!employeeRows) {
    // throw new Error("No employee rows found");
    console.log("No employee rows found");
    return undefined;
  }

  let nextEmployeeRow = employeeRows.find((row) => {
    const worksheetEmployeeId = sanitizeText(
      row.getCell(lookupCol).value?.toString()
    );
    return (
      worksheetEmployeeId !== currentEmployeeId &&
      (conditions
        ? conditions?.every((condition) => {
            const cell = row.getCell(condition.col);
            return condition.condition(cell);
          })
        : true)
    );
  });

  return nextEmployeeRow;
};

// Find next employee row that does not match current employee by name
export const findNextDifferentEmployeeRowByName = ({
  worksheet,
  currentEmployeeRow,
  lastRowNumber,
  lookupCols,
  conditions,
}: {
  worksheet: ExcelJS.Worksheet;
  currentEmployeeRow: ExcelJS.Row;
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
}) => {
  let currentEmployeeLastName = undefined;
  let currentEmployeeFirstName = undefined;
  if ("lastName" in lookupCols && "firstName" in lookupCols) {
    currentEmployeeLastName = sanitizeText(
      currentEmployeeRow.getCell(lookupCols.lastName).value?.toString()
    );
    currentEmployeeFirstName = sanitizeText(
      currentEmployeeRow.getCell(lookupCols.firstName).value?.toString()
    );
  } else if ("name" in lookupCols) {
    const currentEmployeeName = sanitizeText(
      currentEmployeeRow.getCell(lookupCols.name).value?.toString(),
      { removeSpecialChars: false, removeWhitespace: false }
    );

    if (currentEmployeeName) {
      const nameArray = currentEmployeeName.split(lookupCols.delimiter ?? " ");
      if (nameArray.length > 1 && lookupCols.order === "FIRST_NAME LAST_NAME") {
        currentEmployeeFirstName = nameArray[0]?.trim();
        currentEmployeeLastName = nameArray[1]?.trim();
      } else if (
        nameArray.length > 1 &&
        lookupCols.order === "LAST_NAME FIRST_NAME"
      ) {
        currentEmployeeFirstName = nameArray[1]?.trim();
        currentEmployeeLastName = nameArray[0]?.trim();
      }
    }
  }

  const employeeRows = worksheet.getRows(
    currentEmployeeRow.number,
    lastRowNumber
  );
  if (!employeeRows) {
    // throw new Error("No employee rows found");
    console.log("No employee rows found");
    return undefined;
  }

  let nextEmployeeRow = employeeRows.find((row) => {
    let worksheetEmployeeLastName = undefined;
    let worksheetEmployeeFirstName = undefined;
    if ("lastName" in lookupCols && "firstName" in lookupCols) {
      worksheetEmployeeLastName = sanitizeText(
        row.getCell(lookupCols.lastName).value?.toString()
      );
      worksheetEmployeeFirstName = sanitizeText(
        row.getCell(lookupCols.firstName).value?.toString()
      );
    } else if ("name" in lookupCols) {
      const worksheetEmployeeName = sanitizeText(
        row.getCell(lookupCols.name).value?.toString(),
        { removeSpecialChars: false, removeWhitespace: false }
      );
      if (worksheetEmployeeName) {
        const nameArray = worksheetEmployeeName.split(
          lookupCols.delimiter ?? " "
        );
        if (
          nameArray.length > 1 &&
          lookupCols.order === "FIRST_NAME LAST_NAME"
        ) {
          worksheetEmployeeFirstName = nameArray[0]?.trim();
          worksheetEmployeeLastName = nameArray[1]?.trim();
        } else if (
          nameArray.length > 1 &&
          lookupCols.order === "LAST_NAME FIRST_NAME"
        ) {
          worksheetEmployeeFirstName = nameArray[1]?.trim();
          worksheetEmployeeLastName = nameArray[0]?.trim();
        }
      }
    }
    if (!worksheetEmployeeLastName || !worksheetEmployeeFirstName) {
      return true;
    }
    return (
      (worksheetEmployeeLastName !== currentEmployeeLastName ||
        worksheetEmployeeFirstName !== currentEmployeeFirstName) &&
      (conditions
        ? conditions?.every((condition) => {
            const cell = row.getCell(condition.col);
            return condition.condition(cell);
          })
        : true)
    );
  });

  return nextEmployeeRow;
};

// Create new row for employees not found
export const createNewRowForEmployeeNotFound = ({
  worksheet,
  startRowNumber,
  identifierCol,
}: {
  worksheet: ExcelJS.Worksheet;
  startRowNumber: number;
  identifierCol: string;
}) => {
  // Get rows for employee data
  let unknownEmployeePrepRowNumber = startRowNumber;
  let unknownEmployeePopulatedFieldValue = worksheet
    .getRow(unknownEmployeePrepRowNumber)
    .getCell(identifierCol).value;
  while (
    unknownEmployeePopulatedFieldValue !== null &&
    unknownEmployeePopulatedFieldValue !== ""
  ) {
    unknownEmployeePrepRowNumber++;
    unknownEmployeePopulatedFieldValue = worksheet
      .getRow(unknownEmployeePrepRowNumber)
      .getCell(identifierCol).value;
  }
  const unknownEmployeePrepLastEmployeeRow = unknownEmployeePrepRowNumber - 1;

  // Duplicate previous row
  worksheet.duplicateRow(unknownEmployeePrepLastEmployeeRow, 1, true);
  const newRow = worksheet.getRow(unknownEmployeePrepLastEmployeeRow + 1);
  return newRow;
};
