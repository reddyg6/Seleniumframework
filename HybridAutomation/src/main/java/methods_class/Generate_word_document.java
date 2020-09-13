package methods_class;

import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.concurrent.TimeUnit;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.itextpdf.text.Image;

import javax.imageio.ImageIO;

public class Generate_word_document {
	static  XWPFDocument docx =null;
	static XWPFRun run =null;
	static FileOutputStream out =null;
	 	
	 	 public static void initializeDocumen() throws Exception{
		
	 		 
	         
			 	String dirPath = "D:/file" ;
	         
	         	docx = new XWPFDocument();
	         	FileOutputStream out = new FileOutputStream(dirPath+"doc1.docx");
	         	 System.out.println("Write to doc file sucessfully...");
	         
	 	 	}
	             
	            // Add for loop for example, because here we are capturing 5 screenshots
	            
	            public static void writeScreenShotToDocument(ArrayList<byte[]> byteArray, int format) throws Exception{

	            	Image im = null;
	            	for (byte[] bytes : byteArray) {
	            		im = Image.getInstance(bytes);
	            		
	            		docx.addPictureData(bytes, format);
	            		docx.createParagraph();
	            		
	            	}
	           
	            out.flush();
	            out.close();
	            docx.close();
	            docx.write(out);
	            
	            }

	      
	 

	  
	  
	  
	  }
