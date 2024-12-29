package Utilities;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class ExcelUtilities {
    public static FileInputStream fi;
    public static FileOutputStream fo;
    public static XSSFWorkbook wb;
    public static XSSFSheet ws;
    public static XSSFRow row;
    public static XSSFCell cell;
    public static CellStyle style;

    public static String filePath = System.getProperty("user.dir") + "\\test_data\\";

    public static int getRowCount(String xlFile, String xlSheet) throws IOException {

        String path = filePath+xlFile;

        try {
            fi = new FileInputStream(path);
            wb = new XSSFWorkbook(fi);
            ws = wb.getSheet(xlSheet);
            if (ws == null) {
                System.err.println("Sheet \"" + xlSheet + "\" not found in file \"" + xlFile + "\".");
                return -1;
            }
            return ws.getLastRowNum();

        } catch (IOException e) {
            e.printStackTrace();
            return -1;
        }finally {
            wb.close();
            fi.close();
        }

    }

    public static int getCellCount(String xlFile, String xlSheet, int rowNum) throws IOException {

        String path = filePath+xlFile;

        try {
            fi = new FileInputStream(path);
            wb = new XSSFWorkbook(fi);
            ws = wb.getSheet(xlSheet);
            if (ws == null) {
                System.err.println("Sheet \"" + xlSheet + "\" not found in file \"" + xlFile + "\".");
                return -1;
            }
            row = ws.getRow(rowNum);
            return row.getLastCellNum();

        } catch (IOException e) {
            e.printStackTrace();
            return -1;
        }finally {
            wb.close();
            fi.close();
        }

    }

    public static String getCellData(String xlFile, String xlSheet, int rowNum, int colNum) throws IOException {

        String path = filePath+xlFile;

        try{
            fi = new FileInputStream(path);
            wb = new XSSFWorkbook(fi);
            ws = wb.getSheet(xlSheet);
            if (ws == null) {
                System.err.println("Sheet \"" + xlSheet + "\" not found in file \"" + xlFile + "\".");
                return null;
            }
            row = ws.getRow(rowNum);
            cell = row.getCell(colNum);
            // meth 1 : return cell.toString()
            // meth 2 :
            DataFormatter formatter = new DataFormatter();
            return formatter.formatCellValue(cell);

        }catch (IOException e){
            e.printStackTrace();
            return null;
        }finally {
            wb.close();
            fi.close();
        }
    }

    public static void setCellData(String xlFile, String xlSheet, int rowNum, int colNum, String data) throws IOException {

        String path = filePath+xlFile;

        try{
            fi = new FileInputStream(path);
            wb = new XSSFWorkbook(fi);
            ws = wb.getSheet(xlSheet);
            if (ws == null) {
                System.err.println("Sheet \"" + xlSheet + "\" not found in file \"" + xlFile + "\".");
            }
            row = ws.getRow(rowNum);
            cell = row.createCell(colNum);
            cell.setCellValue(data);
            fo = new FileOutputStream(path);
            wb.write(fo);

        }catch (IOException e){
            e.printStackTrace();
        }finally {
            wb.close();
            fi.close();
            fo.close();
        }
    }

    public static void fillGreenColor(String xlFile, String xlSheet, int rowNum , int colNum) throws IOException {

        String path = filePath+xlFile;

        try{

            fi = new FileInputStream(path);
            wb = new XSSFWorkbook(fi);
            ws = wb.getSheet(xlSheet);
            if (ws == null) {
                System.err.println("Sheet \"" + xlSheet + "\" not found in file \"" + xlFile + "\".");
            }
            row = ws.getRow(rowNum);
            cell = row.getCell(colNum);

            style = wb.createCellStyle();
            style.setFillForegroundColor(IndexedColors.GREEN.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            cell.setCellStyle(style);

            fo = new FileOutputStream(path);
            wb.write(fo);

        }catch (IOException e){
            e.printStackTrace();
        }finally {
            wb.close();
            fi.close();
            fo.close();
        }

    }

    public static void fillRedColor(String xlFile, String xlSheet, int rowNum , int colNum) throws IOException {

        String path = filePath+xlFile;

        try{

            fi = new FileInputStream(path);
            wb = new XSSFWorkbook(fi);
            ws = wb.getSheet(xlSheet);
            if (ws == null) {
                System.err.println("Sheet \"" + xlSheet + "\" not found in file \"" + xlFile + "\".");
            }
            row = ws.getRow(rowNum);
            cell = row.getCell(colNum);

            style = wb.createCellStyle();
            style.setFillForegroundColor(IndexedColors.RED.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            cell.setCellStyle(style);

            fo = new FileOutputStream(path);
            wb.write(fo);

        }catch (IOException e){
            e.printStackTrace();
        }finally {
            wb.close();
            fi.close();
            fo.close();
        }

    }


}
