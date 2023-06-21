package reports;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import jxl.Sheet;
import jxl.Workbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import com.google.common.io.Files;
public class Employee_Reports {
	public static WebDriver driver;
	public Workbook wb; 
	public Sheet sh;
	@BeforeTest 	
	public void launchApp()  {					
		System.setProperty("webdriver.chrome.driver", "D:\\ECLIPS\\ORANGE_HRM(TIME)\\ChromeDriver\\chromedriver.exe"); 													
		driver = new ChromeDriver();													
		driver.get("https://adminjon-osondemand.orangehrm.com/symfony/web/index.php/auth/login");													
		driver.manage().window().maximize();													
	}
	@AfterTest		
	public void closeApp() { 		
	driver.close();		
	}		
	@Test						
	public void validationApp(){		
		try {
		FileInputStream f = new FileInputStream("D:\\ECLIPS\\ORANGE_HRM(TIME)\\testfolder\\ORANGEHRM.xls");									
		Workbook wb = Workbook.getWorkbook(f);
		Sheet s = wb.getSheet("Sheet1");			
		driver.findElement(By.name(s.getCell(1, 2).getContents())).sendKeys(s.getCell(1, 5).getContents());											
		driver.findElement(By.name(s.getCell(1, 3).getContents())).sendKeys(s.getCell(1, 6).getContents());											
		driver.findElement(By.name(s.getCell(1, 4).getContents())).click();
		FileInputStream f1 = new FileInputStream("D:\\ECLIPS\\ORANGE_HRM(TIME)\\testfolder\\Reports.xls");	
		try (HSSFWorkbook wb1 = new HSSFWorkbook(f1)) {
			HSSFSheet s1 = wb1.getSheet("Sheet2");
		WebElement time = driver.findElement(By.xpath(s1.getRow(1).getCell(1).getStringCellValue()));	
		Actions act=new Actions(driver);	
		act.moveToElement(time).perform();
		Thread.sleep(3000);	
		WebElement reports = driver.findElement(By.xpath(s1.getRow(2).getCell(1).getStringCellValue()));	
		Actions act1=new Actions(driver); 	
		act1.moveToElement(reports).click().perform();
		Thread.sleep(3000);	
		WebElement empreports = driver.findElement(By.xpath(s1.getRow(3).getCell(1).getStringCellValue()));	
		act1.moveToElement(empreports).click().perform();
		WebElement employee = driver.findElement(By.xpath(s1.getRow(4).getCell(1).getStringCellValue()));	
		employee.sendKeys("Stephen Robert");	
		employee.sendKeys(Keys.ENTER);
		Thread.sleep(3000);
		WebElement pro = driver.findElement(By.xpath(s1.getRow(5).getCell(1).getStringCellValue()));	
		Select sel=new Select(pro);	
		sel.selectByValue("1");
		WebElement acty = driver.findElement(By.xpath(s1.getRow(6).getCell(1).getStringCellValue()));	
		Select sel1=new Select(acty);	
		sel1.selectByValue("-1");
		WebElement projectdaterange = driver.findElement(By.xpath(s1.getRow(7).getCell(1).getStringCellValue()));	
		projectdaterange.click();
		WebElement month = driver.findElement(By.xpath(s1.getRow(8).getCell(1).getStringCellValue()));	
		Select sel2=new Select(month);	
		sel2.selectByValue("2");
		WebElement year = driver.findElement(By.xpath(s1.getRow(9).getCell(1).getStringCellValue()));	
		Select sel3=new Select(year);	
		sel3.selectByValue("2023");
		String searchdate = "25";
		List<WebElement> alldates = driver.findElements(By.xpath(s1.getRow(10).getCell(1).getStringCellValue()));
		for (WebElement elements : alldates)
		{
			String dt=elements.getText();
			
		    if (dt.equals(searchdate))
		    {
		        elements.click(); 
		        break;
		    }
		 		}
	    WebElement projectdaterange1 = driver.findElement(By.xpath(s1.getRow(11).getCell(1).getStringCellValue()));	
		projectdaterange1.click();
		Thread.sleep(3000);	
		WebElement month1 = driver.findElement(By.xpath(s1.getRow(12).getCell(1).getStringCellValue()));	
		Select sel4=new Select(month1);	
		sel4.selectByValue("6");
		WebElement year1 = driver.findElement(By.xpath(s1.getRow(13).getCell(1).getStringCellValue()));	
		Select sel5=new Select(year1);	
		sel5.selectByValue("2023");
		String searchdate1 = "25";
		List<WebElement> alldates1 = driver.findElements(By.xpath(s1.getRow(14).getCell(1).getStringCellValue()));
		for (WebElement elements : alldates1)
		{
			String dt=elements.getText();
			
		    if (dt.equals(searchdate1))
		    {
		        elements.click(); 
		        break;
		    }
		 		}
	    Thread.sleep(3000);	
	    WebElement check = driver.findElement(By.xpath(s1.getRow(15).getCell(1).getStringCellValue()));	
	    check.click();
	    WebElement view = driver.findElement(By.xpath(s1.getRow(16).getCell(1).getStringCellValue()));	
	    view.click();
		}
	    Thread.sleep(10000);
		}
		catch (Exception e) {
		File f11 = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);	
		try {
		Files.copy(f11, new File("D:\\ECLIPS\\ORANGE_HRM(TIME)\\ScreenShots\\employeereports.png"));
		} catch (IOException e1) {
			e1.printStackTrace();
}
}
}
}
