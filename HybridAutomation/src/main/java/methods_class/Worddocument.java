package methods_class;                

import java.awt.Image;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.concurrent.TimeUnit;

import javax.imageio.ImageIO;

import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import com.assertthat.selenium_shutterbug.core.Shutterbug;
import com.assertthat.selenium_shutterbug.utils.web.ScrollStrategy;

import io.github.bonigarcia.wdm.WebDriverManager;

	@SuppressWarnings("unused")			
	public class Worddocument {
		//public static XWPFDocument document;
	static WebDriver driver;
	public static void Generate_word_document() throws Exception {
		
		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		String url="https://www.hsbc.co.uk/insurance/products/home/";
		String url1="http://www.newtours.demoaut.com/";
		driver.get(url1);
		
	      //Write the Document in file system
		  try {
			  XWPFDocument document = new XWPFDocument(); 
			  XWPFParagraph para=document.createParagraph();
			  XWPFRun run=para.createRun();
		
	      run.addPicture(bytearrayconversion(), XWPFDocument.PICTURE_TYPE_JPEG, "", Units.toEMU(500), Units.toEMU(700));
	      
	      driver.get("https://www.tsb.co.uk/personal/");
	      
 
	      run.addPicture(bytearrayconversion(), XWPFDocument.PICTURE_TYPE_JPEG, "", Units.toEMU(400), Units.toEMU(700));
	      
	    
	      FileOutputStream out = new FileOutputStream( new File("F:/createdocument.docx"));
	       
	      
	      document.write(out);
	      out.flush();
	      out.close();
	      document.close();
	      
		  } catch (Exception e) {
				
				e.printStackTrace();
			}
	      System.out.println("createdocument.docx written successully");
	      
	      driver.quit();
	}

	
	
	public static ByteArrayInputStream bytearrayconversion() throws IOException {
		BufferedImage full_bufferimage = Shutterbug.shootPage(driver,ScrollStrategy.WHOLE_PAGE,500,true).getImage();
		ByteArrayOutputStream baos = new ByteArrayOutputStream();
		ImageIO.write(full_bufferimage, "PNG", baos );
		ByteArrayInputStream  data=new ByteArrayInputStream(baos.toByteArray());
		//output.close();
		return data;
		
		}
	
}
