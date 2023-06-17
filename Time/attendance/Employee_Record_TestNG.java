package attendance;									
import java.io.File;
import java.io.FileInputStream;
import java.util.List;

import jxl.Sheet;
import jxl.Workbook;
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
public class Employee_Record_TestNG {
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
		Thread.sleep(3000);	
		act.moveToElement(time).perform();
		WebElement attendance = driver.findElement(By.xpath("//a[@id='menu_attendance_Attendance']"));	
		Actions act1=new Actions(driver); 	
		act1.moveToElement(attendance).click().perform();
		WebElement myrecords = driver.findElement(By.xpath("//*[@id=\"menu_attendance_viewAttendanceRecord\"]"));	
		act1.moveToElement(myrecords).click().perform();
		WebElement emp = driver.findElement(By.xpath("//*[@id=\"attendance_employeeName_empName\"]"));
		emp	.sendKeys("Stephen Robert");	
		emp.sendKeys(Keys.ENTER);
		Thread.sleep(3000);	
		WebElement attendancedate = driver.findElement(By.xpath("//input[@id='attendance_date']"));
		attendancedate.click();
		Thread.sleep(3000);
		WebElement month = driver.findElement(By.xpath("//*[@id=\"ui-datepicker-div\"]/div/div/select[1]"));	
		Select sel=new Select(month);	
		sel.selectByValue("5");	
		WebElement year = driver.findElement(By.xpath("//*[@id=\"ui-datepicker-div\"]/div/div/select[2]"));	
		Select sel1=new Select(year);	
		Thread.sleep(3000);	
		sel1.selectByValue("2023");
		String date = "1";
		List<WebElement> alldates = driver.findElements(By.xpath("//table[@class='ui-datepicker-calendar']//td"));
		for (WebElement elements : alldates)
		{
			String dt=elements.getText();
			
		    if (dt.equals(date))
		    {
		        elements.click(); 
		        break;
		    }
		 		}
		
	
		WebElement view = driver.findElement(By.xpath("//input[@id='btView']"));
		view.click();
		WebElement addattendance = driver.findElement(By.xpath("//*[@id=\"btnPunchOut\"]"));
		addattendance.click();
		
		WebElement in = driver.findElement(By.xpath("//input[@name='button']"));
		in.click();
		WebElement attendancedate1 = driver.findElement(By.xpath("//input[@id='attendance_date']"));
		attendancedate1.click();
		Thread.sleep(3000);
		WebElement month1 = driver.findElement(By.xpath("//*[@id=\"ui-datepicker-div\"]/div/div/select[1]"));	
		Select sel123=new Select(month1);	
		sel123.selectByValue("5");	
		Thread.sleep(3000);	
		WebElement year11 = driver.findElement(By.xpath("//*[@id=\"ui-datepicker-div\"]/div/div/select[2]"));	
		Select sel111=new Select(year11);	
		Thread.sleep(3000);	
		sel111.selectByValue("2023");
		String date1 = "1";
		//Thread.sleep(60000);
		
		List<WebElement> alldates1 = driver.findElements(By.xpath("//table[@class='ui-datepicker-calendar']//td"));
		for (WebElement elements : alldates1)
		{
		String dt=elements.getText();
			
		  if (dt.equals(date1))
		    {
		        elements.click(); 
		        break;
		   }
		 		}
		
		WebElement out = driver.findElement(By.xpath("//input[@name='button']"));
		out.click();
		Thread.sleep(1000);	
		File f1 = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);											
		Files.copy(f1, new File("D:\\ECLIPS\\ORANGE_HRM(TIME)\\ScreenShots\\employeerecords.png"));											

}
}
