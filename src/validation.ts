import ExcelJS from "@zurmokeeper/exceljs";
import { sanitizeText } from "./helpers";

export const validateColumns = (
  headerRow: ExcelJS.Row,
  columnsToBeValidated: { column: string; label: string }[]
) => {
  columnsToBeValidated.map((columnToBeValidated) => {
    const cell = headerRow.getCell(columnToBeValidated.column);
    const cellText = typeof cell.value === "string" ? cell.value : cell.text;
    if (sanitizeText(cellText) !== sanitizeText(columnToBeValidated.label)) {
      throw new Error(
        `TEMPLATE VALIDATION ERROR: Expected column ${
          columnToBeValidated.label === ""
            ? "(No label)"
            : columnToBeValidated.label
        } in column ${columnToBeValidated.column} but found ${cellText}`
      );
    }
  });
};
