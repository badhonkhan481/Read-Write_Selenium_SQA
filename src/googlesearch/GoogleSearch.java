package googlesearch;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import java.util.concurrent.TimeUnit;
import org.openqa.selenium.Keys;
//import static jdk.nashorn.internal.objects.NativeObject.keys;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import java.io.IOException;
import java.time.Duration;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GoogleSearch{
    
   public static void main(String[] args) throws InterruptedException, FileNotFoundException, IOException {
   System.setProperty("webdriver.chrome.driver","C:\\Users\\badhonkhan481\\Desktop\\SQA Project\\chromedriver_win32\\chromedriver.exe");
      WebDriver driver = new ChromeDriver();
      driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
      driver.manage().window().maximize();
      driver.navigate().to("https://www.google.com/");

      FileInputStream fis = new FileInputStream("C:\\Users\\badhonkhan481\\Desktop\\SQA Project\\GoogleSearch\\Excel.xlsx");
//     File src=new File("C:\\Users\\badhonkhan481\\Documents\\NetBeansProjects\\javawrite\\Excel.xlsx");
//    FileInputStream fis = new FileInputStream(src);
//      XSSFWorkbook workbook = new XSSFWorkbook(fis);
      XSSFWorkbook workbook = new XSSFWorkbook(fis);
      XSSFSheet sheet = workbook.getSheet("Saturday");
      int rowcount =sheet.getLastRowNum();
      int colcount=sheet.getRow(2).getLastCellNum();
      System.out.println("rowcount:"+rowcount+"colconut:"+colcount);
      
      
      
      for(int i=2; i<=rowcount; ++i)
      {
        XSSFRow celldata=sheet.getRow(i);
        String key=celldata.getCell(2).getStringCellValue();
        XSSFRow Cell=sheet.getRow(i);
        Row row = sheet.getRow(i);
	Cell cell = row.createCell(3);
        Cell cell2 = row.createCell(4);
        WebElement p=driver.findElement(By.name("q"));
        p.sendKeys(key);
        WebDriverWait w = new WebDriverWait(driver,Duration.ofSeconds(10));
        w.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//ul")));
        Thread.sleep(100);  
//        WebElement ele=driver.findElement(By.xpath("//ul"));
        

        WebElement ele = driver.findElement(By.xpath("//*[@id=\"Zrbbw\"]"));
        WebElement ele2 = driver.findElement(By.xpath("//*[@id=\"vTtioc\"]"));
        

        String eleText=getText(ele);
        String eleText2=getText2(ele2);
        
        System.out.println("Text value:"+eleText);
	cell.setCellValue(eleText);
        cell2.setCellValue(eleText2);
   
        p.clear();
  
//        p.submit();
      }
      
     FileOutputStream fos = new FileOutputStream("C:\\Users\\badhonkhan481\\Desktop\\SQA Project\\GoogleSearch\\Excel.xlsx");
	workbook.write(fos);
	fos.close();
	System.out.println("END OF WRITING DATA IN EXCEL");
   
   }
   
   public static String getText(WebElement ele){
        String eleText=null;
        eleText=ele.getText();
        eleText="";
        if(eleText.equals("")){
            eleText=ele.getAttribute("innerText");
            eleText="";
            if(eleText.equals("")){
                eleText=ele.getAttribute("textContent");
            }
        }
        return eleText;
   }
  public static String getText2(WebElement ele2){
        String eleText2=null;
        eleText2=ele2.getText();
        eleText2="";
        if(eleText2.equals("")){
            eleText2=ele2.getAttribute("innerText");
            eleText2="";
            if(eleText2.equals("")){
                eleText2=ele2.getAttribute("textContent");
            }
        }
        return eleText2;
   }
 
}
