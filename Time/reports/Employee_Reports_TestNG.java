package reports;
import java.io.File;
import java.io.FileInputStream;
import java.util.List;

import jxl.Sheet;
import jxl.Workbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
//import org.openqa.selenium.OutputType;
//import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.google.common.io.Files;
public class Employee_Reports_TestNG {
	public static WebDriver driver;
	public Workbook wb; 
	public Sheet sh;
	@BeforeTest 	
	public void LaunchApp() throws Exception, Throwable {					
		System.setProperty("webdriver.chrome.driver", "D:\\ECLIPS\\ORANGE_HRM(TIME)\\ChromeDriver\\chromedriver.exe"); 													
		driver = new ChromeDriver();													
		driver.get("https://adminjon-osondemand.orangehrm.com/symfony/web/index.php/auth/login");													
		driver.manage().window().maximize();													
	}
	@AfterTest		
	public void CloseApp() { 		
	driver.close();		
	}		
	@Test						
	public void ValidationApp() throws Throwable, Exception {						
		FileInputStream f = new FileInputStream("D:\\ECLIPS\\ORANGE_HRM(TIME)\\testfolder\\ORANGEHRM.xls");									
		Workbook wb = Workbook.getWorkbook(f);
		Sheet s = wb.getSheet("Sheet1");			
		driver.findElement(By.name(s.getCell(1, 2).getContents())).sendKeys(s.getCell(1, 5).getContents());											
		driver.findElement(By.name(s.getCell(1, 3).getContents())).sendKeys(s.getCell(1, 6).getContents());											
		driver.findElement(By.name(s.getCell(1, 4).getContents())).click();
		WebElement time = driver.findElement(By.xpath("//*[@id=\"menu_time_viewTimeModule\"]/b"));	
		Actions act=new Actions(driver);	
		act.moveToElement(time).perform();
		Thread.sleep(3000);	
		WebElement reports = driver.findElement(By.xpath("//*[@id=\"menu_time_Reports\"]"));	
		Actions act1=new Actions(driver); 	
		act1.moveToElement(reports).click().perform();
		Thread.sleep(3000);	
		WebElement empreports = driver.findElement(By.xpath("//*[@id=\"menu_time_displayEmployeeReportCriteria\"]"));	
		act1.moveToElement(empreports).click().perform();
		WebElement employee = driver.findElement(By.xpath("//*[@id=\"employee_empName\"]"));	
		employee.sendKeys("Stephen Robert");	
		employee.sendKeys(Keys.ENTER);
		Thread.sleep(3000);
		WebElement pro = driver.findElement(By.xpath("//*[@id=\"time_project_name\"]"));	
		Select sel=new Select(pro);	
		sel.selectByValue("1");
		WebElement acty = driver.findElement(By.xpath("//*[@id=\"time_activity_name\"]"));	
		Select sel1=new Select(acty);	
		sel1.selectByValue("-1");
		WebElement projectdaterange = driver.findElement(By.xpath("//*[@id=\"project_date_range_from_date\"]"));	
		projectdaterange.click();
		WebElement month = driver.findElement(By.xpath("//*[@id=\"ui-datepicker-div\"]/div/div/select[1]"));	
		Select sel2=new Select(month);	
		sel2.selectByValue("2");
		WebElement year = driver.findElement(By.xpath("//*[@id=\"ui-datepicker-div\"]/div/div/select[2]"));	
		Select sel3=new Select(year);	
		sel3.selectByValue("2023");
		String searchdate = "25";
		List<WebElement> alldates = driver.findElements(By.xpath("//table[@class='ui-datepicker-calendar']//td"));
		for (WebElement elements : alldates)
		{
			String dt=elements.getText();
			
		    if (dt.equals(searchdate))
		    {
		        elements.click(); 
		        break;
		    }
		 		}
	    WebElement projectdaterange1 = driver.findElement(By.xpath("//*[@id=\"project_date_range_to_date\"]"));	
		projectdaterange1.click();
		Thread.sleep(3000);	
		WebElement month1 = driver.findElement(By.xpath("//*[@id=\"ui-datepicker-div\"]/div/div/select[1]"));	
		Select sel4=new Select(month1);	
		sel4.selectByValue("6");
		WebElement year1 = driver.findElement(By.xpath("//*[@id=\"ui-datepicker-div\"]/div/div/select[2]"));	
		Select sel5=new Select(year1);	
		sel5.selectByValue("2023");
		String searchdate1 = "25";
		List<WebElement> alldates1 = driver.findElements(By.xpath("//table[@class='ui-datepicker-calendar']//td"));
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
	    WebElement check = driver.findElement(By.xpath("//*[@id=\"only_include_approved_timesheets\"]"));	
	    check.click();
	    WebElement view = driver.findElement(By.xpath("//*[@id=\"viewbutton\"]"));	
	    view.click();
	    Thread.sleep(10000);	
		File f1 = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);											
		Files.copy(f1, new File("D:\\ECLIPS\\ORANGE_HRM(TIME)\\ScreenShots\\employeereports.png"));													
}
}

