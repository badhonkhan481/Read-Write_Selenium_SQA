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
   System.setProperty("webdriver.chrome.driver","C:\\Users\\badhonkhan481\\Desktop\\SQL Project\\chromedriver_win32\\chromedriver.exe");
      WebDriver driver = new ChromeDriver();
      driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
      driver.manage().window().maximize();
      driver.navigate().to("https://www.google.com/");

      FileInputStream fs = new FileInputStream("C:\\Users\\badhonkhan481\\Desktop\\SQL Project\\GoogleSearch\\Excel.xlsx");
      XSSFWorkbook workbook = new XSSFWorkbook(fs);
      XSSFSheet sheet = workbook.getSheet("Saturday");
      int rowcount =sheet.getLastRowNum();
      int colcount=sheet.getRow(2).getLastCellNum();
      System.out.println("rowcount:"+rowcount+"colconut:"+colcount);
      
      
      for(int i=2; i<=rowcount; ++i)
      {
        XSSFRow celldata=sheet.getRow(i);
        String key=celldata.getCell(2).getStringCellValue();
        WebElement p=driver.findElement(By.name("q"));
        p.sendKeys(key);
        WebDriverWait w = new WebDriverWait(driver,Duration.ofSeconds(10));
        w.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//ul")));
        Thread.sleep(100);  
//        WebElement ele=driver.findElement(By.xpath("//ul"));
        
        WebElement ele = driver.findElement(By.xpath("//*[@id=\"APjFqb\"]"));

        String eleText=getText(ele);
        System.out.println("Text value:"+eleText);
//        GoogleSearch we=new GoogleSearch();
//       String excelPath="C:\\Users\\badhonkhan481\\Desktop\\SQL Project\\GoogleSearch\\Excel.xlsx";
//       we.writeData(excelPath, "Saturday", 2, 3,eleText);
      
         
        p.clear();
        
        
        
//        p.submit();
      }
//      GoogleSearch we=new GoogleSearch();
//       String excelPath="C:\\Users\\badhonkhan481\\Desktop\\SQL Project\\GoogleSearch\\Excel.xlsx";
//       we.writeData(excelPath, "Saturday", 2, 3,"eleText");
//         
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
  
//   public void writeData(String excelPath, String sheetName, int rowNumber, int columnNumber, String data) {
//   try{
//       File file=new File(excelPath);
//      FileInputStream fis =new FileInputStream(file);
//       XSSFWorkbook wb =new XSSFWorkbook();
//       XSSFSheet sheet =wb.getSheet(sheetName);
//       XSSFRow row =sheet.getRow(rowNumber);
//       XSSFCell cell=row.getCell(columnNumber );
////       Row.MissingCellPolicy RETURN_BLANK_AS_NULL = org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL;
//       if(cell==null){
//           row.createCell(columnNumber);
//           cell.setCellValue(data);
//               }
//       else{
//               cell.setCellValue(data);    
//                   }
//
//       FileOutputStream fio=new FileOutputStream(file);
//       wb.write(fio);
//       wb.close();
//   }catch(IOException io){
//       io.printStackTrace();
//       
//   } 
//   

}
