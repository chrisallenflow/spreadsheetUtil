package com.arcanepost.service.spreadsheetutil;

import java.io.BufferedInputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import au.com.bytecode.opencsv.CSVReader;

/**
 * <p>
 * This utility was build to do two things. First, it was meant to convert either an XLS file or a CSV file to an XLSX
 * format. Finally, it has convenience methods to help parse thru an XLSX file to read data using Apache POI.
 * </p>
 *
 * @author callen
 * @since 01.25.2016
 */
public final class ApachePOIUtil {

    private static final Logger LOGGER = LoggerFactory.getLogger(ApachePOIUtil.class);

    /**
     * Private constructor.
     */
    private ApachePOIUtil() {
    }

    /**
     * Pass in a spreadsheet row, get returned a list of values (as String) in the order that they appear.
     *
     * @param row Spreadsheet row.
     * @return A list of row values.
     * @throws SpreadsheetUtilException Some exception.
     */
    public static ArrayList<String> getXlsxSheetRowValues(Row row) throws SpreadsheetUtilException {

        ArrayList<String> valueList = new ArrayList<>();

        try {
            for (int cn = 0; cn < row.getLastCellNum(); cn++) {
                // If the cell is missing from the file, generate a blank one
                // (Works by specifying a MissingCellPolicy)

                Cell cell = row.getCell(cn, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                String cellValue = "";

                // convert boolean and numeric values to basic string

                switch (cell.getCellTypeEnum()) {

                    case BOOLEAN:
                        cellValue = String.valueOf(cell.getBooleanCellValue());
                        break;
                    case NUMERIC:
                        // check to see if it is a date
                        if (HSSFDateUtil.isCellDateFormatted(cell)) {
                            DataFormatter formatter = new DataFormatter(Locale.US);
                            cellValue = formatter.formatCellValue(cell);
                            break;
                        } else {
                            cellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
                            break;
                        }
                    case STRING:
                        cellValue = cell.getStringCellValue();
                        break;
                    case BLANK:
                        cellValue = "";
                        break;
                    case FORMULA:
                        if (cell.getCachedFormulaResultTypeEnum() == CellType.NUMERIC) {
                            cellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
                        } else if (cell.getCachedFormulaResultTypeEnum() == CellType.STRING) {
                            cellValue = cell.getRichStringCellValue().getString();
                        }
                        break;
                    default:
                        LOGGER.debug("spreadsheet had a cell of type 'ERROR'.  Ignoring.");
                        break;
                }
                valueList.add(cellValue);
            }
        } catch (Exception e) {
            throw new SpreadsheetUtilException(e);
        }
        return valueList;
    }

    /**
     * <p>
     * This method checks the given worksheet. It checks to see that:
     * </p>
     *
     * <ol>
     * <li>The first row are column headers (have column names)</li>
     * <li>All column names passed in with columnList exist in the first row, in the order they appear in the list.</li>
     * <li>The number of columns in the worksheet has to be columnList.size()</li>
     * <li>The second row must have data (not an empty worksheet)</li>
     * </ol>
     *
     * @param workbook         The XLSX workbook.
     * @param sheetNumberIndex A workbook has many worksheets. Identify the sheet number here. First sheet would be "0".
     * @param columnList       A list of String representing the names of the columns
     * @return A list of error messages, or an empty list if no errors.
     * @throws SpreadsheetUtilException Some exception.
     */
    public static ArrayList<String> validateXlsxSheet(XSSFWorkbook workbook, int sheetNumberIndex,
                                                      ArrayList<String> columnList) throws SpreadsheetUtilException {

        ArrayList<String> errorList = new ArrayList<>();
        try {
            // STEP 1: CHECK VALIDATION CRITERIA PASSED IN

            // make sure that there is validation criteria
            if (columnList == null || columnList.size() == 0) {
                errorList.add("No validation criteria given.");
                return errorList;
            }

            // VALIDATION CRITERIA IS GOOD

            // STEP 2: VALIDATE THE WORKSHEET

            // get worksheet to test
            XSSFSheet sheet = workbook.getSheetAt(sheetNumberIndex);

            // Iterate through each row from the worksheet
            Iterator<Row> rowIterator = sheet.iterator();
            if (!rowIterator.hasNext()) {
                errorList.add("Worksheet is empty.");
                return errorList;
            }
            boolean firstRow = true;
            int currentRow = 0;
            while (rowIterator.hasNext()) {
                if (firstRow) {
                    // look for proper number of columns, by name
                    Row row = rowIterator.next();
                    currentRow++;
                    int columnCounter = 0;
                    // For each row, iterate through each columns
                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {
                        columnCounter++;
                        Cell cell = cellIterator.next();

                        String cellValue = "";

                        // convert boolean and numeric values to basic string

                        switch (cell.getCellTypeEnum()) {

                            case BOOLEAN:
                                cellValue = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case NUMERIC:
                                cellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
                                break;
                            case STRING:
                                cellValue = cell.getStringCellValue();
                                break;
                            case BLANK:
                                cellValue = "";
                                break;
                            case FORMULA:
                                if (cell.getCachedFormulaResultTypeEnum() == CellType.NUMERIC) {
                                    cellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
                                } else if (cell.getCachedFormulaResultTypeEnum() == CellType.STRING) {
                                    cellValue = cell.getRichStringCellValue().getString();
                                }
                                break;
                            default:
                                break;
                        }

                        if (columnCounter <= columnList.size()
                                && !columnList.get(columnCounter - 1).equalsIgnoreCase(cellValue.trim())) {
                            errorList.add("Column " + columnCounter + " must be '" + columnList.get(columnCounter - 1)
                                    + "', but it is '" + cellValue.trim() + "'");
                        }
                    }
                    if (columnCounter != columnList.size()) {
                        errorList.add("Incorrect number of columns. Should be " + columnList.size() + ", actual is "
                                + columnCounter);
                    }

                    firstRow = false;
                } else {
                    // Row row = rowIterator.next();
                    rowIterator.next();
                    currentRow++;
                }
            }
            if (currentRow < 2) {
                errorList.add("Empty worksheet.  No data after columns.");
            }
        } catch (Exception e) {
            throw new SpreadsheetUtilException(e);
        }
        return errorList;
    }

    /**
     * Converts an Excel spreadsheet InputStream that is in the "xls" format to one that is in the "xlsx" format.
     *
     * @param fis InputStream of an Excel Spreadsheet in xls format.
     * @return InputStream of an Excel Spreadsheet in the xlsx format.
     * @throws SpreadsheetUtilException Some exception.
     */
    public static InputStream convertXlsToXlsx(InputStream fis) throws SpreadsheetUtilException {

        InputStream in = new BufferedInputStream(fis);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        Workbook wbIn = null;
        Workbook wbOut = null;
        try {
            wbIn = new HSSFWorkbook(in);

            wbOut = new XSSFWorkbook();
            int sheetCnt = wbIn.getNumberOfSheets();
            LOGGER.debug("Number of sheets = '" + sheetCnt + "'");
            for (int i = 0; i < sheetCnt; i++) {
                Sheet sIn = wbIn.getSheetAt(i);
                LOGGER.debug("sheet in = '" + sIn.getSheetName() + "'");
                Sheet sOut = wbOut.createSheet(sIn.getSheetName());
                Iterator<Row> rowIt = sIn.rowIterator();
                while (rowIt.hasNext()) {
                    Row rowIn = rowIt.next();
                    Row rowOut = sOut.createRow(rowIn.getRowNum());

                    Iterator<Cell> cellIt = rowIn.cellIterator();
                    while (cellIt.hasNext()) {
                        Cell cellIn = cellIt.next();
                        Cell cellOut = rowOut.createCell(cellIn.getColumnIndex(), cellIn.getCellTypeEnum());
                        switch (cellIn.getCellTypeEnum()) {

                            case BOOLEAN:
                                cellOut.setCellValue(String.valueOf(cellIn.getBooleanCellValue()));
                                break;
                            case NUMERIC:
                                cellOut.setCellValue(NumberToTextConverter.toText(cellIn.getNumericCellValue()));
                                break;
                            case STRING:
                                cellOut.setCellValue(cellIn.getStringCellValue());
                                break;
                            case BLANK:
                                cellOut.setCellValue("");
                                break;
                            case ERROR:
                                cellOut.setCellValue(cellIn.getErrorCellValue());
                                break;
                            case FORMULA:
                                if (cellIn.getCachedFormulaResultTypeEnum() == CellType.NUMERIC) {
                                    cellOut.setCellValue(NumberToTextConverter.toText(cellIn.getNumericCellValue()));
                                } else if (cellIn.getCachedFormulaResultTypeEnum() == CellType.STRING) {
                                    cellOut.setCellValue(cellIn.getRichStringCellValue().getString());
                                }
                                break;
                            default:
                                break;

                        }

                        CellStyle styleIn = cellIn.getCellStyle();
                        CellStyle styleOut = cellOut.getCellStyle();
                        styleOut.setDataFormat(styleIn.getDataFormat());

                        cellOut.setCellComment(cellIn.getCellComment());

                    }
                }
            }

            wbOut.write(out);
            return new ByteArrayInputStream(out.toByteArray());

        } catch (Exception e) {
            throw new SpreadsheetUtilException(e);
        } finally {
            // try to close streams, if they are not null
            try {
                in.close();
            } catch (Exception e) {
                LOGGER.info("Issue closing Inputstream. Ignoring...");
            }
            try {
                out.close();
            } catch (Exception e) {
                LOGGER.info("Issue closing Outputstream. Ignoring...");
            }
            try {
                assert wbIn != null;
                wbIn.close();
            } catch (Exception e) {
                LOGGER.info("Issue closing Workbook (wbIn).  Ignoring...");
            }
            try {
                assert wbOut != null;
                wbOut.close();
            } catch (Exception e) {
                LOGGER.info("Issue closing Workbook (wbOut).  Ignoring...");
            }
        }
    }

    /**
     * Converts a CSV file to an InputStream of an Excel Spreadsheet in the xlsx format.
     *
     * @param in File as an InputStream
     * @return InputStream of an Excel Spreadsheet in the xlsx format.
     * @throws SpreadsheetUtilException Some exception.
     */
    public static InputStream convertCsvToXlsx(InputStream in) throws SpreadsheetUtilException {

        CSVReader reader = null;
        XSSFWorkbook workBook = null;
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        try {
            List<String[]> myEntries;

            // reader = new CSVReader(new FileReader(file));
            reader = new CSVReader(new InputStreamReader(in));
            myEntries = reader.readAll();

            workBook = new XSSFWorkbook();
            XSSFSheet sheet = workBook.createSheet("sheet1");
            int rownum = 0;

            if (myEntries != null) {
                for (String[] rowArray : myEntries) {
                    rownum++;
                    XSSFRow currentRow = sheet.createRow(rownum);
                    for (int i = 0; i < rowArray.length; i++) {
                        currentRow.createCell(i).setCellValue(rowArray[i]);
                    }
                }
            }

            workBook.write(out);
            return new ByteArrayInputStream(out.toByteArray());

        } catch (Exception e) {
            throw new SpreadsheetUtilException(e);
        } finally {
            // try to close streams, if they are not null
            try {
                out.close();
            } catch (Exception e) {
                LOGGER.info("Issue closing OutputStream.  Ignoring...");
            }
            try {
                assert workBook != null;
                workBook.close();
            } catch (Exception e) {
                LOGGER.info("Issue closing Workbook.  Ignoring...");
            }
            try {
                reader.close();
            } catch (Exception e) {
                LOGGER.info("Issue closing Reader.  Ignoring...");
            }
        }
    }

    /**
     * Private method to convert a double to a String in a clean manner.
     *
     * @param d the double.
     * @return the converted String.
     * @deprecated Use NumberToTextConverter.toText() instead.
     */
    @Deprecated
    private static String convertDoubleToString(double d) {
        if (d == (long) d) {
            return String.format("%d", (long) d);
        } else {
            return String.format("%s", d);
        }
    }

    /**
     * Convert XLSX file to a JSON Object. No validation.
     *
     * @param workbook An xlsx workbook.
     * @return Spreadsheet as a JSON object.
     * @throws SpreadsheetUtilException Some exception.
     */
    public static JSONObject convertXlsxToJson(XSSFWorkbook workbook) throws SpreadsheetUtilException {

        JSONObject obj = new JSONObject();
        try {
            int rownumber = 0;
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            if (!rowIterator.hasNext()) {
                LOGGER.debug("Violation:  Worksheet is empty.");
            } else {
                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    rownumber++;
                    JSONArray list = new JSONArray();
                    ArrayList<String> rowValueList = getXlsxSheetRowValues(row);
                    int longestLength = 0;
                    for (String cellData : rowValueList) {
                        if (cellData.trim().length() > longestLength) {
                            longestLength = cellData.trim().length();
                        }
                        list.put(cellData.trim());
                    }
                    if (longestLength > 0) {
                        obj.put(rownumber + "", list);
                    }
                }
            }
        } catch (Exception e) {
            throw new SpreadsheetUtilException(e);
        }
        return obj;
    }

}
