package HybridAutomation;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.Color;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Start {
	public static WebDriver driver;
   //public static Robot rb;
	public static void main(String[] args) throws AWTException, InterruptedException {
		// TODO Auto-generated method stub
	
		
				System.setProperty("webdriver.chrome.driver", "F://chromedriver.exe");
				 driver= new ChromeDriver();
				 driver.manage().window().maximize();
				 driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
				 driver.get("https://insuranceportal.av.qs.online-insure.com/youshaped/Buildings/Assumptions");
			
				 
			// driver.findElement(By.tagName("body")).sendKeys(Keys.F12);
				 
			List<WebElement> ele=driver.findElements(By.tagName("h2"));

			System.out.println(driver.findElement(By.tagName("h1")));
			//(new Capturescreenshot()).pagegetpagesource();
			//System.out.println(driver.getPageSource());
			//driver.quit();
			WebDriverWait wait= new WebDriverWait(driver,5);
			wait.until(ExpectedConditions.presenceOfElementLocated(By.tagName("h2")));
			
			int totalsize=ele.size();
			for(int i=0;i<=totalsize-1;i++){
				
			
				//System.out.println(i);
				if(ele.get(i).isDisplayed()==true){
				
				 //capture web Element screen shot
				/* File fl=ele.get(i).getScreenshotAs(OutputType.FILE);
				FileUtils.copyFile(fl, new File("F:\\"+ele.get(i).getText()+".png"));*/
					
					
					 
					/*  //capture screen shot of the web element 
					 File screenshot=((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
					
					Point p=ele.get(i).getLocation();
					int width=ele.get(i).getSize().getWidth();
					int height=ele.get(i).getSize().getHeight();
					
					Rectangle rect = new Rectangle(width, height);
					
					 BufferedImage img=null;
					 img= ImageIO.read(screenshot);
					 
					 BufferedImage destination=img.getSubimage(p.getX(), p.getY(), rect.width, rect.height);
					 
					 ImageIO.write(destination, "png", screenshot);
					    FileUtils.copyFile(screenshot,
				                new File("F:\\"+ele.get(i).getText()+"_headingh3.png"));*/
					
				   
					 
				}
			}
			
			//Robot class to handle keyboard events/functions
				
		
		  Robot rb= new Robot(); rb.keyPress(KeyEvent.VK_F12);
		  rb.keyRelease(KeyEvent.VK_F12);
		 
			//robot_press_f12();
			
				wait(3);
				
		
		  rb.keyPress(KeyEvent.VK_CONTROL); rb.keyPress(KeyEvent.VK_SHIFT);
		  rb.keyPress(KeyEvent.VK_D);
		  rb.keyRelease(KeyEvent.VK_D); rb.keyRelease(KeyEvent.VK_SHIFT);
		  rb.keyRelease(KeyEvent.VK_CONTROL);
		 
				// robot_dockside_bottom();
				
				wait(3);
				
			//	robot_inspector_element();
		
		  rb.keyPress(KeyEvent.VK_CONTROL); rb.keyPress(KeyEvent.VK_SHIFT);
		  rb.keyPress(KeyEvent.VK_C); rb.keyRelease(KeyEvent.VK_CONTROL);
		  rb.keyRelease(KeyEvent.VK_SHIFT); rb.keyRelease(KeyEvent.VK_C);
		 

				wait(3);
			
			for(int i=0; i<=totalsize-1;i++){
			
			wait(2);
			
			if(ele.get(i).isDisplayed()) {
			
			
				String str_font=ele.get(i).getCssValue("font-family");      

				System.out.println(str_font);
				
				String str_color=ele.get(i).getCssValue("color");
				
				String hexacolor=Color.fromString(str_color).asHex();
				
				System.out.println(hexacolor);
				
			/*Actions act =new Actions(driver);
			act.clickAndHold(ele.get(i));
			
			
			System.out.println("mousehover on heading");
			wait(3);
			
			File src= ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
			FileUtils.copyFile(src, new File("F://"+ele.get(i).getText()+".jpeg"));
			
			Rectangle screenRect = new Rectangle(Toolkit.getDefaultToolkit().getScreenSize());
			BufferedImage capture = new Robot().createScreenCapture(screenRect);
			ImageIO.write(capture, "png", new File("F://"+ele.get(i).getText()+".png"));
			*/
			}  }
			
			//TakesScreenshot take= ( (TakesScreenshot) driver);
			
			wait(3);
			driver.quit();
			}
			
		public static void wait( int time) throws InterruptedException{
			Thread.sleep(time*1000);
			
		}
		
		public static void robot_press_f12() throws AWTException {
			Robot rb= new Robot();
			rb.keyPress(KeyEvent.VK_F12);
			rb.keyRelease(KeyEvent.VK_F12);
		}
		
		public static void robot_dockside_bottom() throws AWTException {
			Robot rb= new Robot();
			rb.keyPress(KeyEvent.VK_CONTROL);
			rb.keyPress(KeyEvent.VK_SHIFT);
			rb.keyPress(KeyEvent.VK_D);
			
			rb.keyRelease(KeyEvent.VK_D);
			rb.keyRelease(KeyEvent.VK_SHIFT);
			rb.keyRelease(KeyEvent.VK_CONTROL);
		}
		
		public static void robot_inspector_element() throws AWTException {
			Robot rb= new Robot();
			rb.keyPress(KeyEvent.VK_CONTROL);
			rb.keyPress(KeyEvent.VK_SHIFT);
			rb.keyPress(KeyEvent.VK_C);
			
			rb.keyRelease(KeyEvent.VK_C);
			rb.keyRelease(KeyEvent.VK_SHIFT);
			rb.keyRelease(KeyEvent.VK_CONTROL);
		}
		
		}

		
	


