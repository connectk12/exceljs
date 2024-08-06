import ExcelJS from "@zurmokeeper/exceljs";
import { sanitizeText } from './helpers';

export const validateColumns = (
  headerRow: ExcelJS.Row,
  columnsToBeValidated: { column: string; label: string }[],
) => {
  columnsToBeValidated.map((columnToBeValidated) => {
    const cell = headerRow.getCell(columnToBeValidated.column);
    if (
      sanitizeText(cell.value?.toString()) !==
      sanitizeText(columnToBeValidated.label)
    ) {
      throw new Error(
        `TEMPLATE VALIDATION ERROR: Expected column ${columnToBeValidated.label === "" ? "(No label)" : columnToBeValidated.label} in column ${columnToBeValidated.column} but found ${cell.value?.toString()}`,
      );
    }
  });
};