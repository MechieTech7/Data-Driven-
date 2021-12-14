package excelLaunch;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

public class ReadExcel {
    public void sheetRead() throws IOException {

        WebDriverManager.chromedriver().setup();
        WebDriver driver = new ChromeDriver();

        FileInputStream inputStream = new FileInputStream("C:\\Users\\lokesh.chandramurthy\\IdeaProjects\\AutomatingExcel\\src\\main\\java\\excelLaunch\\excelPath.properties");
        Properties properties = new Properties();
        properties.load(inputStream);
        String path = properties.getProperty("filePath");


        File file = new File(path);
        FileInputStream input = new FileInputStream(file);
        XSSFWorkbook wb = new XSSFWorkbook(input);
        // XSSFSheet sheet=wb.getSheet("demo.xlsx");
        XSSFSheet sheet = wb.getSheetAt(0);

        int rowCount = sheet.getPhysicalNumberOfRows();
        //  int rowCount=sheet.getLastRowNum()-sheet.getFirstRowNum();

        driver.get("https://demoqa.com/automation-practice-form");


        WebElement firstName = driver.findElement(By.id("firstName"));
        WebElement lastName = driver.findElement(By.id("lastName"));
        WebElement email = driver.findElement(By.id("userEmail"));
        WebElement genderMale = driver.findElement(By.id("gender-radio-1"));
        WebElement mobile = driver.findElement(By.id("userNumber"));
        WebElement address = driver.findElement(By.id("currentAddress"));
        WebElement submitBtn = driver.findElement(By.id("submit"));


        for (int i = 1; i < rowCount; i++) {
            firstName.sendKeys(sheet.getRow(i).getCell(0).getStringCellValue());
            lastName.sendKeys(sheet.getRow(i).getCell(1).getStringCellValue());
            email.sendKeys(sheet.getRow(i).getCell(2).getStringCellValue());
            DataFormatter dataFormatter = new DataFormatter();
            String a = dataFormatter.formatCellValue( sheet.getRow(i).getCell(4));
            mobile.sendKeys(a);
            address.sendKeys(sheet.getRow(i).getCell(5).getStringCellValue());

            JavascriptExecutor js = (JavascriptExecutor) driver;
            js.executeScript("arguments[0].click();", genderMale);

            js.executeScript("arguments[0].click();", submitBtn);



            WebElement confirmationMessage = driver.findElement(By.xpath("//div[text()='Thanks for submitting the form']"));



            if (confirmationMessage.isDisplayed()) {
                System.out.println("All values passed");

            } else {
                System.out.println("Fail!");;
            }

            WebElement closebtn = driver.findElement(By.id("closeLargeModal"));
            js.executeScript("arguments[0].click();", closebtn);


            driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
            // driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));
        }


        wb.close();
        driver.quit();
    }


    public static void main(String[] args) throws IOException {
        ReadExcel callMethod = new ReadExcel();
        callMethod.sheetRead();
    }
}
