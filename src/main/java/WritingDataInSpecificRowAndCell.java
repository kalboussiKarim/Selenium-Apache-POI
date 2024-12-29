import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

public class WritingDataInSpecificRowAndCell {

    public static void main(String[] args) {
        //Excel file ----> Workbook ----> Sheets ----> Rows ----> Cells

        try {
            // 1 - opening the file
            FileOutputStream file = new FileOutputStream( System.getProperty("user.dir")+"\\test_data\\SpecificRow&CellData.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("Data");

            XSSFRow row1 = sheet.createRow(3);
            row1.createCell(4).setCellValue("Test...");

            workbook.write(file);
            workbook.close();
            file.close();


        }catch (Exception e){
            System.out.println(e.toString());
        }

    }
}
