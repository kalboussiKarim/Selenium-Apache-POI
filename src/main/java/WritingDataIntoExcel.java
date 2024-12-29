import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class WritingDataIntoExcel {

    public static void main(String[] args) {
        //Excel file ----> Workbook ----> Sheets ----> Rows ----> Cells

        try {
            // 1 - opening the file
            FileOutputStream file = new FileOutputStream( System.getProperty("user.dir")+"\\test_data\\genratedData.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("Data");

            XSSFRow row1 = sheet.createRow(0);
            row1.createCell(0).setCellValue("Java");
            row1.createCell(1).setCellValue(1234);
            row1.createCell(2).setCellValue("Automation");

            XSSFRow row2 = sheet.createRow(1);
            row2.createCell(0).setCellValue("Python");
            row2.createCell(1).setCellValue(234);
            row2.createCell(2).setCellValue("Automation");

            XSSFRow row3 = sheet.createRow(2);
            row3.createCell(0).setCellValue("PHP");
            row3.createCell(1).setCellValue(2);
            row3.createCell(2).setCellValue("Automation");

            workbook.write(file);
            workbook.close();
            file.close();


        }catch (Exception e){
            System.out.println(e.toString());
        }

    }
}
