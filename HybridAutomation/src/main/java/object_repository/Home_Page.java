package object_repository;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.*;

public class Home_Page {

	@FindBy(id="")
	public WebElement abc;
	
	@FindBy(how = How.ID, using="element ID")
	public WebElement rest; 

	
}
