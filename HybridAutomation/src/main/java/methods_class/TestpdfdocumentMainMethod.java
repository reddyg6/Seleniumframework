package methods_class;

import java.awt.AWTException;
import java.awt.HeadlessException;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;
import org.openqa.selenium.support.Color;
import javax.imageio.ImageIO;
import javax.imageio.stream.ImageOutputStream;

import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

import com.assertthat.selenium_shutterbug.core.PageSnapshot;
import com.assertthat.selenium_shutterbug.core.Shutterbug;
import com.assertthat.selenium_shutterbug.utils.web.ScrollStrategy;
import com.itextpdf.text.Image;

import io.github.bonigarcia.wdm.WebDriverManager;
import ru.yandex.qatools.ashot.AShot;
import ru.yandex.qatools.ashot.Screenshot;
import ru.yandex.qatools.ashot.shooting.ShootingStrategies;

public class TestpdfdocumentMainMethod extends GeneratePDFMethods {

	static WebDriver driver;
	static String Data;
	static int row;

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		String Sheetname = "DataSheet";

		ArrayList<byte[]> screenshotsList = newScreenShotslist();
		ArrayList<String[]> ListCsvData = new ArrayList<String[]>();
		
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		driver.get("https://www.barclays.co.uk/insurance/home-insurance/");
		
		RobotPressF12MoveDoc();
		
		boolean isPdfHaveAtleastObepage = false;
		//Element 1
		List<WebElement> H2 = driver.findElements(By.tagName("h2"));

		
		for (row = 1; row < H2.size(); row++) {
			int i = row - 1;
			if (H2.get(i).isDisplayed() == true) {

				Actions ac = new Actions(driver);
				ac.moveToElement(H2.get(i)).build().perform();

				String h2_color_value = H2.get(i).getCssValue("color");
				String hexcolor = Color.fromString(h2_color_value).asHex();
				boolean color = hexcolor.equalsIgnoreCase("#00395D");

				if (color) {
					isPdfHaveAtleastObepage = true;
					screenshotsList.add(bytearraydata());
				} else {

					System.out.println("Failed");
				}

			}
		}

		if (isPdfHaveAtleastObepage) {
			GeneratePDFMethods.initializeDocument("H2");
			GeneratePDFMethods.writeScreenShotToDocument(screenshotsList);
		}
		
		screenshotsList.clear();
		newScreenShotslist().clear();
		
		System.out.println("Element 1 screenshots doument saved");
		//Element 2 
		List<WebElement> H3 = driver.findElements(By.tagName("h3"));

		for (row = 1; row < H3.size(); row++) {
			int i = row - 1;
			System.out.println(i);

			if (H3.get(i).isDisplayed() == true) {

				Actions ac = new Actions(driver);
				ac.moveToElement(H3.get(i)).build().perform();
				
				String h3_color_value = H3.get(i).getCssValue("color");
				String hexcolor = Color.fromString(h3_color_value).asHex();
				boolean color = hexcolor.equalsIgnoreCase("#00395D");

				if (color) {
					isPdfHaveAtleastObepage = true;
					screenshotsList.add(bytearraydata());
				} else {

					System.out.println("Failed");
				}

			}
		}

		if (isPdfHaveAtleastObepage) {
			GeneratePDFMethods.initializeDocument("H3");
			GeneratePDFMethods.writeScreenShotToDocument(screenshotsList);
		}
		
		System.out.println("Element 2 screenshots doument saved");
		newScreenShotslist().clear();
		//System.out.println();
		
		driver.quit();
	}

	
	

}
