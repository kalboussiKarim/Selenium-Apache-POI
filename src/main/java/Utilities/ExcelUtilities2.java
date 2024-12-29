package Utilities;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;

public class ExcelUtilities2 {

    private static final String FILE_PATH = System.getProperty("user.dir") + "\\test_data\\";


    private static XSSFWorkbook openWorkbook(String path) throws IOException {
        try (FileInputStream fi = new FileInputStream(path)) {
            return new XSSFWorkbook(fi);
        }
    }


    private static void saveWorkbook(XSSFWorkbook wb, String path) throws IOException {
        try (FileOutputStream fo = new FileOutputStream(path)) {
            wb.write(fo);
        }
    }

    public static int getRowCount(String xlFile, String xlSheet) throws IOException {
        String path = FILE_PATH + xlFile;

        try (XSSFWorkbook wb = openWorkbook(path)) {
            XSSFSheet ws = wb.getSheet(xlSheet);
            if (ws == null) {
                throw new IllegalArgumentException("Sheet \"" + xlSheet + "\" not found in file \"" + xlFile + "\".");
            }
            return ws.getLastRowNum();
        }
    }

    public static int getCellCount(String xlFile, String xlSheet, int rowNum) throws IOException {
        String path = FILE_PATH + xlFile;

        try (XSSFWorkbook wb = openWorkbook(path)) {
            XSSFSheet ws = wb.getSheet(xlSheet);
            if (ws == null) {
                throw new IllegalArgumentException("Sheet \"" + xlSheet + "\" not found in file \"" + xlFile + "\".");
            }
            XSSFRow row = ws.getRow(rowNum);
            return row != null ? row.getLastCellNum() : -1;
        }
    }

    public static String getCellData(String xlFile, String xlSheet, int rowNum, int colNum) throws IOException {
        String path = FILE_PATH + xlFile;

        try (XSSFWorkbook wb = openWorkbook(path)) {
            XSSFSheet ws = wb.getSheet(xlSheet);
            if (ws == null) {
                throw new IllegalArgumentException("Sheet \"" + xlSheet + "\" not found in file \"" + xlFile + "\".");
            }
            XSSFRow row = ws.getRow(rowNum);
            if (row == null) return null;

            XSSFCell cell = row.getCell(colNum);
            if (cell == null) return null;

            DataFormatter formatter = new DataFormatter();
            return formatter.formatCellValue(cell);
        }
    }

    public static void setCellData(String xlFile, String xlSheet, int rowNum, int colNum, String data) throws IOException {
        String path = FILE_PATH + xlFile;

        try (XSSFWorkbook wb = openWorkbook(path)) {
            XSSFSheet ws = wb.getSheet(xlSheet);
            if (ws == null) {
                throw new IllegalArgumentException("Sheet \"" + xlSheet + "\" not found in file \"" + xlFile + "\".");
            }
            XSSFRow row = ws.getRow(rowNum);
            if (row == null) row = ws.createRow(rowNum);

            XSSFCell cell = row.createCell(colNum);
            cell.setCellValue(data);

            saveWorkbook(wb, path);
        }
    }

    public static void fillCellColor(String xlFile, String xlSheet, int rowNum, int colNum, IndexedColors color) throws IOException {
        String path = FILE_PATH + xlFile;

        try (XSSFWorkbook wb = openWorkbook(path)) {
            XSSFSheet ws = wb.getSheet(xlSheet);
            if (ws == null) {
                throw new IllegalArgumentException("Sheet \"" + xlSheet + "\" not found in file \"" + xlFile + "\".");
            }
            XSSFRow row = ws.getRow(rowNum);
            if (row == null) row = ws.createRow(rowNum);

            XSSFCell cell = row.getCell(colNum);
            if (cell == null) cell = row.createCell(colNum);

            CellStyle style = wb.createCellStyle();
            style.setFillForegroundColor(color.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(style);

            saveWorkbook(wb, path);
        }
    }

    public static void fillGreenColor(String xlFile, String xlSheet, int rowNum, int colNum) throws IOException {
        fillCellColor(xlFile, xlSheet, rowNum, colNum, IndexedColors.GREEN);
    }

    public static void fillRedColor(String xlFile, String xlSheet, int rowNum, int colNum) throws IOException {
        fillCellColor(xlFile, xlSheet, rowNum, colNum, IndexedColors.RED);
    }
}
