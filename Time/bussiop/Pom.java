package bussiop;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
public class Pom {
	@FindBy(name="txtUsername") 
	public WebElement uid;
	@FindBy(name="txtPassword")
	public WebElement password;
	@FindBy(name="Submit") 
	public WebElement sub;

}
