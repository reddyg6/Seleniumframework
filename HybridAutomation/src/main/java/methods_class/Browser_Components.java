package methods_class;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.ie.*;
import org.openqa.selenium.chrome.*;
import org.openqa.selenium.firefox.*;
public class Browser_Components 
{
	static WebDriver driver;
	static String Gecko_driver_property_path;
	static String IE_driver_property_path;
	static String chrome_driver_property_path;
       public static void launch_browser(String browser){
    	   
    	   switch (browser) {
		case "Firefox":
			System.setProperty("webdriver.gecko.driver", Gecko_driver_property_path);
			driver=new FirefoxDriver();
			break;

		case "IE":
			System.setProperty("wwebdriver.ie.driver",IE_driver_property_path);
			driver=new InternetExplorerDriver();
			break;
		case "Chrome":
			System.setProperty("webdriver.chrome.driver", chrome_driver_property_path);
			driver= new ChromeDriver();
			
    	   }
	
    	
        
       }
    
}
