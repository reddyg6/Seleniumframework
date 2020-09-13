package methods_class;

import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.Color;

import io.github.bonigarcia.wdm.WebDriverManager;

public class WriteDataIntoExcel {
	static WebDriver driver;
	static String Data;
	static int row;
	static XSSFWorkbook workbook =null;
	 static XSSFSheet sheet=null;
	static Map<String, Object[]> data = null;
	public static void main(String[] args) {
		
		data=new TreeMap<String, Object[]>();
		//Blank workbook
        workbook = new XSSFWorkbook(); 
         
        //Create a blank sheet
        sheet = workbook.createSheet("Employee Data");
          
		/*WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		driver.get("https://www.barclays.co.uk/insurance/home-insurance/");
		
		/*
		 * List<WebElement> H2 = driver.findElements(By.tagName("h2")); data.put("1",
		 * new Object[] {"ID", "NAME", "LASTNAME"}); for (row = 1; row < H2.size();
		 * row++) { int i = row - 1; if (H2.get(i).isDisplayed() == true) {
		 * 
		 * Actions ac = new Actions(driver);
		 * ac.moveToElement(H2.get(i)).build().perform();
		 * 
		 * String h2_color_value = H2.get(i).getCssValue("color"); String hexcolor =
		 * Color.fromString(h2_color_value).asHex(); boolean color =
		 * hexcolor.equalsIgnoreCase("#00395D");
		 * 
		 * data.put(null, new Object[] {row, hexcolor, color}); } }
		 */
		        
		        //This data needs to be written (Object[])
		       
		          
		        //Iterate over data and write to sheet
		        Set<String> keyset = data.keySet();
		        int rownum = 0;
		        for (String key : keyset)
		        {
		            Row row = sheet.createRow(rownum++);
		            Object [] objArr = data.get(row);
		            int cellnum = 0;
		            for (Object obj : objArr)
		            {
		               org.apache.poi.ss.usermodel.Cell cell = row.createCell(cellnum++);
		               if(obj instanceof String)
		                    cell.setCellValue((String)obj);
		                else if(obj instanceof Integer)
		                    cell.setCellValue((Integer)obj);
		            }
		        }
		        
		        try
		        
		        {
		            //Write the workbook in file system
		            FileOutputStream out = new FileOutputStream("D:/testdata.xlsx");
		            workbook.write(out);
		            out.close();
		            System.out.println("Excel file saved");
		        } 
		        catch (Exception e) 
		        {
		            e.printStackTrace();
		        }
		    	

      	}
 
    }
