import Utilities.ExcelUtilities;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import java.io.IOException;
import java.time.Duration;

public class FDCalculator {

    public static void main(String[] args) throws IOException {

        WebDriver driver = new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(10));
        driver.get("https://www.moneycontrol.com/fixed-income/calculator/state-bank-of-india-sbi/fixed-deposit-calculator-SBI-BSB001.html");
        driver.manage().window().maximize();

        String fileName = "caldata.xlsx";
        String sheetName = "Sheet1";

        int rows = ExcelUtilities.getRowCount(fileName,sheetName);

        for (int i = 1; i < rows; i++) {
            // reading data from the Excel file
            String principle = ExcelUtilities.getCellData(fileName,sheetName,i,0);
            String rateOfInterest = ExcelUtilities.getCellData(fileName,sheetName,i,1);
            String periodDuration = ExcelUtilities.getCellData(fileName,sheetName,i,2);
            String periodUnit = ExcelUtilities.getCellData(fileName,sheetName,i,3);
            String frequency = ExcelUtilities.getCellData(fileName,sheetName,i,4);
            String maturityValue = ExcelUtilities.getCellData(fileName,sheetName,i,5);

            // pass data into application
            driver.findElement(By.xpath("//input[@id='principal']")).sendKeys(principle);
            driver.findElement(By.xpath("//input[@id='interest']")).sendKeys(rateOfInterest);
            driver.findElement(By.xpath("//input[@id='tenure']")).sendKeys(periodDuration);

            Select selectPeriodUnit =  new Select(driver.findElement(By.xpath("//select[@id='tenurePeriod']")));
            selectPeriodUnit.selectByVisibleText(periodUnit);

            Select selectFrequency =  new Select(driver.findElement(By.xpath("//select[@id='frequency']")));
            selectFrequency.selectByVisibleText(frequency);

            driver.findElement(By.xpath("//img[@src='https://images.moneycontrol.com/images/mf_revamp/btn_calcutate.gif']")).click();
            String result = driver.findElement(By.xpath("//div[@id='resp_prin']/following-sibling::div/span[2]/strong")).getText();

            System.out.println(result);

            if(Double.parseDouble(result)==Double.parseDouble(maturityValue)){
                System.out.println("Test passed");
                ExcelUtilities.setCellData(fileName,sheetName,i,7,"Passed");
                ExcelUtilities.fillGreenColor(fileName,sheetName,i,7);
            }else {
                System.out.println("Test failed");
                ExcelUtilities.setCellData(fileName,sheetName,i,7,"Failed");
                ExcelUtilities.fillRedColor(fileName,sheetName,i,7);
            }

            driver.findElement(By.xpath("//img[@class='PL5']")).click();
        }
        driver.quit();
    }
}
