package excelLaunch;

import io.github.bonigarcia.wdm.WebDriverManager;
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
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.concurrent.TimeUnit;

public class ReadExcel {
    public  void sheetRead() throws IOException {

        WebDriverManager.chromedriver().setup();
        WebDriver driver= new ChromeDriver();


        File file =    new File("C:\\Users\\lokesh.chandramurthy\\IdeaProjects\\AutomatingExcel\\src\\main\\demo.xlsx");
        FileInputStream inputStream = new FileInputStream(file);
        XSSFWorkbook wb=new XSSFWorkbook(inputStream);
       // XSSFSheet sheet=wb.getSheet("demo.xlsx");
        XSSFSheet sheet=wb.getSheetAt(0);

        int rowCount=sheet.getPhysicalNumberOfRows();
      //  int rowCount=sheet.getLastRowNum()-sheet.getFirstRowNum();

        driver.get("https://demoqa.com/automation-practice-form");


        WebElement firstName=driver.findElement(By.id("firstName"));
        WebElement lastName=driver.findElement(By.id("lastName"));
        WebElement email=driver.findElement(By.id("userEmail"));
        WebElement genderMale= driver.findElement(By.id("gender-radio-1"));
        WebElement mobile=driver.findElement(By.id("userNumber"));
        WebElement address=driver.findElement(By.id("currentAddress"));
        WebElement submitBtn=driver.findElement(By.id("submit"));


        for(int i=1;i<=rowCount;i++) {
            firstName.sendKeys(sheet.getRow(i).getCell(0).getStringCellValue());
            lastName.sendKeys(sheet.getRow(i).getCell(1).getStringCellValue());
            email.sendKeys(sheet.getRow(i).getCell(2).getStringCellValue());
            mobile.sendKeys(sheet.getRow(i).getCell(4).getStringCellValue());
            address.sendKeys(sheet.getRow(i).getCell(5).getStringCellValue());

            JavascriptExecutor js = (JavascriptExecutor) driver;
            js.executeScript("arguments[0].click();", genderMale);


            submitBtn.click();

            WebElement confirmationMessage = driver.findElement(By.xpath("//div[text()='Thanks for submitting the form']"));


          
            FileOutputStream outputStream = new FileOutputStream("C:\\Users\\lokesh.chandramurthy\\IdeaProjects\\AutomatingExcel\\src\\main\\demo.xlsx");
            wb.write(outputStream);


            WebElement closebtn = driver.findElement(By.id("closeLargeModal"));
            closebtn.click();
            driver.manage().timeouts().implicitlyWait(2000, TimeUnit.SECONDS);
            //driver.manage().timeouts().implicitlyWait(Duration.of(5,"Seconds"));
        }


        wb.close();
        driver.quit();
    }

    public static void main(String[] args) throws IOException {
        ReadExcel callMethod = new ReadExcel();
        callMethod.sheetRead();
    }
}
