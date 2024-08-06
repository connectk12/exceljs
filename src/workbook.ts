import ExcelJS from "@zurmokeeper/exceljs";

/**
 * Open an Excel workbook from a file at the designated workbookPathname
 * Default behavior: sets modified Date to current date, ignores conditional formatting nodes, and calculates all formulas on load.
 * Supports password-protected workbooks.

 * @param workbookPathname - Pathname of the workbook file
 * @param opts - Options
 * @param opts.date - Date to set the workbook's modified date to
 * @param opts.password - Password to decrypt the workbook
 * @param opts.ignoreNodes - Nodes to ignore when reading the workbook
 * @returns The Excel workbook
 */
export const openWorkbook = async ({
  workbookPathname,
  opts,
}: {
  workbookPathname: string;
  opts?: {
    date?: Date;
    password?: string;
    ignoreNodes?: string[];
  };
}) => {
  const workbook = new ExcelJS.Workbook();
  workbook.modified = opts?.date ?? new Date();
  await workbook.xlsx.readFile(workbookPathname, {
    password: opts?.password,
    ignoreNodes: opts?.ignoreNodes ?? ["conditionalFormatting"],
  });
  workbook.calcProperties.fullCalcOnLoad = true;
  console.log("Workbook opened", workbookPathname);
  return workbook;
};

/**
 * Exports Excel workbook to a designated outputPathname
 * Default behavior: sets modified Date to current date.
 *
 * @param workbook - The Excel workbook to export
 * @param outputPathname - Pathname to export the workbook to
 * @param opts - Options
 * @param opts.worksheetName - Name of the worksheet to set the view to
 * @param opts.currentWorksheet - Current worksheet to remove shared formulas from
 * @param opts.removeSharedFormulas - Whether to remove shared formulas from the current worksheet
 * @param opts.setWorksheetViewId - Whether to set the worksheet view to the worksheet with the given name
 * @param opts.password - Password to encrypt the workbook
 * @returns The Excel workbook
 */
export const exportWorkbook = async ({
  workbook,
  outputPathname,
  opts,
}: {
  workbook: ExcelJS.Workbook;
  outputPathname: string;
  opts?: {
    worksheetName?: string;
    currentWorksheet?: ExcelJS.Worksheet;
    removeSharedFormulas?: boolean;
    setWorksheetViewId?: boolean;
    password?: string;
  };
}) => {
  if (opts?.currentWorksheet && opts.removeSharedFormulas) {
    removeSharedFormulas(opts.currentWorksheet);
  }
  if (opts?.currentWorksheet && opts.setWorksheetViewId && opts.worksheetName) {
    let worksheetId = -1;
    workbook.worksheets.forEach((worksheet, index) => {
      if (worksheet.name === opts.worksheetName) {
        worksheetId = index + 1;
      }
    });
    if (worksheetId > 0) {
      setWorksheetView(workbook, worksheetId);
    }
  }
  await workbook.xlsx.writeFile(outputPathname);
  console.log("Workbook exported", outputPathname);
};

export const removeSharedFormulas = (worksheet: ExcelJS.Worksheet) => {
  worksheet.eachRow((row) => {
    row.eachCell((cell) => {
      if (cell.formulaType.toString().localeCompare("shared") === 0) {
        cell.value = null;
      }
    });
  });
};

export const setWorksheetView = (workbook: ExcelJS.Workbook, tabId: number) => {
  workbook.views = [
    {
      x: 0,
      y: 0,
      width: 40000,
      height: 20000,
      firstSheet: 0,
      activeTab: tabId - 1,
      visibility: "visible",
    },
  ];
};
