package methods_class;
import java.awt.AWTException;
import java.awt.HeadlessException;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;

import javax.imageio.ImageIO;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;

import com.assertthat.selenium_shutterbug.core.Shutterbug;
import com.assertthat.selenium_shutterbug.utils.web.ScrollStrategy;
import com.itextpdf.text.Document;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;

class GeneratePDFMethods {
static WebDriver driver;
	 static Document document = null;
	 
		static PdfWriter writer ;
		
		 public static void initializeDocument(String Path) { 
			FileOutputStream fos;
			try {
				
				document=new Document(PageSize.A4,20,20,20,20);
				fos = new FileOutputStream(new File("F:\\" + Path + ".pdf"));
				writer = PdfWriter.getInstance(document, fos);
				
			} catch (Exception e) {
			
				e.printStackTrace();
			}
	}
		 
		public static void writeScreenShotToDocument(ArrayList<byte[]> byteArray) throws Exception {
			System.out.println("writeScreenShotToDocument started"+byteArray.size());
			
			//open the document
			document.open();
			
			Image im = null;
			
			for (byte[] bytes : byteArray) {
				
			/*
			 * if (!newScreenShotslist().contains(bytes)) {
			 * 
			 * newScreenShotslist().add(bytes); }
			 */
				im = Image.getInstance(bytes);
				//set the size of the image
				im.scaleToFit(PageSize.A4.getWidth(), PageSize.A4.getHeight());
				// add the image to PDF
				document.add(im);
				document.add(new Paragraph(""));
				System.out.println("added");
			}
			
			document.close();
			writer.close();
		}

		public static ArrayList<byte[]> newScreenShotslist() {
			ArrayList<byte[]> screenshotsList = new ArrayList<byte[]>();
			return screenshotsList;
		}
		//screenshots.add(bytearraydata());

		public static byte[] bytearrayconversion() throws IOException {
			BufferedImage bufferimage = Shutterbug.shootPage(TestpdfdocumentMainMethod.driver,ScrollStrategy.WHOLE_PAGE,500,true).getImage();
			ByteArrayOutputStream output = new ByteArrayOutputStream();
			ImageIO.write(bufferimage, "PNG", output );
			byte[] data=output.toByteArray();
			return data;
			}
		public static void WaitTimeinSeconds(int number) throws InterruptedException {
			Thread.sleep(1000*number);
		}
		
		public static void RobotPressF12MoveDoc() throws AWTException, InterruptedException {
			
			Robot rb= new Robot();
			rb.keyPress(KeyEvent.VK_F12);
			WaitTimeinSeconds(2);
			rb.keyPress(KeyEvent.VK_CONTROL);
			rb.keyPress(KeyEvent.VK_SHIFT);
			rb.keyPress(KeyEvent.VK_D);
			WaitTimeinSeconds(2);
			rb.keyRelease(KeyEvent.VK_CONTROL);
			rb.keyRelease(KeyEvent.VK_SHIFT);
			rb.keyRelease(KeyEvent.VK_D);
			WaitTimeinSeconds(2);
			rb.keyPress(KeyEvent.VK_CONTROL);
			rb.keyPress(KeyEvent.VK_SHIFT);
			rb.keyPress(KeyEvent.VK_C);
			WaitTimeinSeconds(2);
			rb.keyRelease(KeyEvent.VK_CONTROL);
			rb.keyRelease(KeyEvent.VK_SHIFT);
			rb.keyRelease(KeyEvent.VK_C);
			WaitTimeinSeconds(2);
			rb.keyRelease(KeyEvent.VK_F12);
		}
		
		
		public static byte[] bytearraydata() throws HeadlessException, AWTException, IOException {
			BufferedImage screenFullImage = new Robot().createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
			ByteArrayOutputStream output = new ByteArrayOutputStream();
			ImageIO.write(screenFullImage, "PNG", output );
			byte[] inout=output.toByteArray();
			return inout;
		}
		

}
