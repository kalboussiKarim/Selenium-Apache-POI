import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.util.Scanner;

public class WritingDynamicDataIntoExcel {

    public static void main(String[] args) {
        //Excel file ----> Workbook ----> Sheets ----> Rows ----> Cells

        try {
            // 1 - opening the file
            FileOutputStream file = new FileOutputStream( System.getProperty("user.dir")+"\\test_data\\DynamicData.xlsx");
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("DynamicData");

            Scanner scanner = new Scanner(System.in);

            System.out.println("How many rows you want to enter ? :");
            int noRows = scanner.nextInt();

            System.out.println("How many cells you want to enter ? :");
            int noCells = scanner.nextInt();

            for (int r = 0; r < noRows; r++) {
                System.out.println("You are currently filling row no "+r+" :");
                XSSFRow row = sheet.createRow(r);
                for (int c = 0; c < noCells; c++) {
                    System.out.println("Enter cell no: "+c+" value :");
                    row.createCell(c).setCellValue(scanner.next());
                }
            }

            workbook.write(file);
            workbook.close();
            file.close();


        }catch (Exception e){
            System.out.println(e.toString());
        }

    }
}
