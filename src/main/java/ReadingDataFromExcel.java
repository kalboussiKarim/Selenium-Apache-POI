import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

public class ReadingDataFromExcel {

    public static void main(String[] args) {
        //Excel file ----> Workbook ----> Sheets ----> Rows ----> Cells


        try {
            // 1 - opening the file
            FileInputStream file = new FileInputStream( System.getProperty("user.dir")+"\\test_data\\data1.xlsx");

            // 2 - Workbook
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // 3 - Sheet
            XSSFSheet sheet = workbook.getSheet("Sheet1");
            //XSSFSheet sheet = workbook.getSheetAt(0);

            // 4 - Rows & Cells
                // first we need fetch the last row number because it represents the number of rows
            int totalRows = sheet.getLastRowNum();
                // second we need fetch the number of cells
            int totalCells = sheet.getRow(1).getLastCellNum(); // we can pout any row number 0; 1 ...

            System.out.println("Number of Rows is : "+ totalRows);
            System.out.println("Number of Cells is : "+ totalCells);

            for (int r = 0; r<= totalRows; r++){
                XSSFRow currentRow = sheet.getRow(r);
                for (int c = 0;c<totalCells;c++){
                    XSSFCell currentCell =  currentRow.getCell(c);
                    System.out.print(currentCell.toString()+ "\t");
                }
                System.out.println(" ");

            }

            workbook.close();
            file.close();

        }catch (Exception e){
            System.out.println(e.toString());
        }

    }
}