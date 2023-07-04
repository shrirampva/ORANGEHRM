package reports;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

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
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import com.google.common.io.Files;
import bussiop.Pom;
public class Attendance_Summary {
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
	public void validationApp() {				
		try {
		FileInputStream f = new FileInputStream("D:\\ECLIPS\\ORANGE_HRM(TIME)\\testfolder\\ORANGEHRM.xls");									
		Workbook wb = Workbook.getWorkbook(f);
		Sheet s = wb.getSheet("Sheet1");			
		Pom p =PageFactory.initElements(driver,Pom.class);
		p.uid.sendKeys(s.getCell(1, 5).getContents());
		p.password.sendKeys(s.getCell(1, 6).getContents());
		p.sub.click();
		FileInputStream f1 = new FileInputStream("D:\\ECLIPS\\ORANGE_HRM(TIME)\\testfolder\\Reports.xls");	
		try (HSSFWorkbook wb1 = new HSSFWorkbook(f1)) {
			HSSFSheet s1 = wb1.getSheet("Sheet3");
		WebElement time = driver.findElement(By.xpath(s1.getRow(1).getCell(1).getStringCellValue()));	
		Actions act=new Actions(driver);	
		act.moveToElement(time).perform();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		WebElement reports = driver.findElement(By.xpath(s1.getRow(2).getCell(1).getStringCellValue()));	
		Actions act1=new Actions(driver); 	
		act1.moveToElement(reports).click().perform();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		WebElement AtdSumreports = driver.findElement(By.xpath(s1.getRow(3).getCell(1).getStringCellValue()));	
		act1.moveToElement(AtdSumreports).click().perform();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		WebElement employee = driver.findElement(By.xpath(s1.getRow(4).getCell(1).getStringCellValue()));	
		employee.sendKeys("Stephen Robert");	
		employee.sendKeys(Keys.ENTER);
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		WebElement jobtlt = driver.findElement(By.xpath(s1.getRow(5).getCell(1).getStringCellValue()));	
		Select sel=new Select(jobtlt);	
		sel.selectByValue("1");
		WebElement subunit = driver.findElement(By.xpath(s1.getRow(6).getCell(1).getStringCellValue()));	
		Select sel1=new Select(subunit);	
		sel1.selectByValue("0");
		WebElement emlyntstus = driver.findElement(By.xpath(s1.getRow(7).getCell(1).getStringCellValue()));	
		Select sel2=new Select(emlyntstus );	
		sel2.selectByValue("1");
		WebElement projectdaterange = driver.findElement(By.xpath(s1.getRow(8).getCell(1).getStringCellValue()));	
		projectdaterange.click();
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		WebElement month = driver.findElement(By.xpath(s1.getRow(9).getCell(1).getStringCellValue()));	
		Select sel3=new Select(month);	
		sel3.selectByValue("2");
		WebElement year = driver.findElement(By.xpath(s1.getRow(10).getCell(1).getStringCellValue()));	
		Select sel4=new Select(year);	
		sel4.selectByValue("2023");
		String searchdate = "25";
		List<WebElement> alldates = driver.findElements(By.xpath(s1.getRow(11).getCell(1).getStringCellValue()));
		for (WebElement elements : alldates)
		{
			String dt=elements.getText();
			
		    if (dt.equals(searchdate))
		    {
		        elements.click(); 
		        break;
		    }
		 		}
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
	    WebElement projectdaterange1 = driver.findElement(By.xpath(s1.getRow(12).getCell(1).getStringCellValue()));	
		projectdaterange1.click();
		WebElement month1 = driver.findElement(By.xpath(s1.getRow(13).getCell(1).getStringCellValue()));	
		Select sel5=new Select(month1);	
		sel5.selectByValue("6");
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);	
		WebElement year1 = driver.findElement(By.xpath(s1.getRow(14).getCell(1).getStringCellValue()));	
		Select sel6=new Select(year1);	
		sel6.selectByValue("2023");
		String searchdate1 = "25";
		List<WebElement> alldates1 = driver.findElements(By.xpath(s1.getRow(15).getCell(1).getStringCellValue()));
		for (WebElement elements : alldates1)
		{
			String dt=elements.getText();
			
		    if (dt.equals(searchdate1))
		    {
		        elements.click(); 
		        break;
		    }
		 		}
	
	    WebElement view = driver.findElement(By.xpath(s1.getRow(16).getCell(1).getStringCellValue()));	
	    view.click();
		}
		driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
		}
		catch (Exception e) {
		File f11 = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);	
		try {
		Files.copy(f11, new File("D:\\ECLIPS\\ORANGE_HRM(TIME)\\ScreenShots\\attendnacesummary.png"));	
		} catch (IOException e1) {
			e1.printStackTrace();
}
		}
}
}
