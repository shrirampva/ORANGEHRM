package bussiop;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
public class Pom {
	@FindBy(name="txtUsername") 
	private WebElement uid;
	@FindBy(name="txtPassword")
	private WebElement password;
	@FindBy(name="Submit") 
	private WebElement sub;
	public WebElement getUid() {
		return uid;
	}
	public void setUid(WebElement uid) {
		this.uid = uid;
	}
	public WebElement getPassword() {
		return password;
	}
	public void setPassword(WebElement password) {
		this.password = password;
	}
	public WebElement getSub() {
		return sub;
	}
	public void setSub(WebElement sub) {
		this.sub = sub;
	}

}
