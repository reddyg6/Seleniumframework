package methods_class;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.Color;

import io.github.bonigarcia.wdm.WebDriverManager;

public class WriteDataInExcel {
	static WebDriver driver;
	static String Data;
	static int row;

	public static void main(String[] args) {
		try {
			System.out.print("Current Time in milliseconds = ");
			long startTime = System.nanoTime();

			WebDriverManager.chromedriver().setup();
			driver = new ChromeDriver();
			driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
			driver.manage().window().maximize();
			driver.get("https://www.barclays.co.uk/insurance/home-insurance/");

			// String[] header = new String[] { "H2_Color", "H2_Color_status",
			// "H2_Font_Family", "Details" };
			int rowCount = 0;

			/*
			 * Workbook workbook = new XSSFWorkbook(); Sheet sheet =
			 * workbook.createSheet("sheet1");
			 */
			// String Excelsheetname="sheet";

			/*
			 * Row row; int columnCount;
			 */
			List<String> header = Arrays.asList("H1_Color", "H1_Color_status", "H1_Font_Family", "Details");
			rowCount = writeHeader(header, rowCount);
			List<WebElement> H1 = driver.findElements(By.tagName("h1"));
			for (int i = 0; i < H1.size(); i++) {
				if (H1.get(i).isDisplayed() == true) {
					System.out.println("h1 callled---");
					String color_value = H1.get(i).getCssValue("color");
					String h1_hexcolor = Color.fromString(color_value).asHex();
					boolean color = h1_hexcolor.equalsIgnoreCase("#00395D");

					String h1_Fontfamily = H1.get(i).getCssValue("font-family").replace(",", "");
					boolean fs = h1_Fontfamily
							.equalsIgnoreCase("\"Expert Sans Light\", \"Trebuchet MS\", Arial, Verdana, sans-serif");
					String[] data = new String[] { h1_hexcolor, "" + color + "", h1_Fontfamily, "" + fs + "" };

					rowCount = writeDatatoExcel(data, rowCount);
				}
			}

			List<String> header1 = Arrays.asList("H2_Color", "H2_Color_status", "H2_Font_Family", "Details");
			rowCount = writeHeader(header1, rowCount);
			List<WebElement> H2 = driver.findElements(By.tagName("h2"));
			for (int i = 0; i < H2.size(); i++) {
				if (H2.get(i).isDisplayed() == true) {
					System.out.println("h2 called----");
					String h2_color_value = H2.get(i).getCssValue("color");
					String h2_hexcolor = Color.fromString(h2_color_value).asHex();
					boolean color = h2_hexcolor.equalsIgnoreCase("#00395D");

			  		String h2_Fontfamily = H2.get(i).getCssValue("font-family").replace(",", "");
					boolean fs = h2_Fontfamily
							.equalsIgnoreCase("\"Expert Sans Light\", \"Trebuchet MS\", Arial, Verdana, sans-serif");
					String[] data = new String[] { h2_hexcolor, "" + color + "", h2_Fontfamily, "" + fs + "" };

					rowCount = writeDatatoExcel(data, rowCount);
				}
			}

			List<WebElement> H3 = driver.findElements(By.tagName("h3"));
			List<String> header2 = Arrays.asList("H3_Color", "H3_Color_status", "H3_Font_Family");

			rowCount = writeHeader(header2, rowCount);
			
			for (int i = 0; i < H3.size(); i++) {
				if (H3.get(i).isDisplayed() == true) {
					String h3_color_value = H3.get(i).getCssValue("color");
					String h3_hexcolor = Color.fromString(h3_color_value).asHex();
					boolean color = h3_hexcolor.equalsIgnoreCase("#00395D");

					String h3_Fontfamily = H3.get(i).getCssValue("font-family");

					String[] data = new String[] { h3_hexcolor, "" + color + "", h3_Fontfamily };
					rowCount = writeDatatoExcel(data, rowCount);
				}
			}

			writeAllData("D:/BarclaysHomePage");

			driver.quit();

		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	static Workbook workbook = new XSSFWorkbook();
	static Sheet sheet = workbook.createSheet("sheet");

	private static int writeDatatoExcel(String[] data, int rowCount) {
		// true -> green
		CellStyle t_style = workbook.createCellStyle();
		CellStyle f_style = workbook.createCellStyle();
		Font truefont = workbook.createFont();
		truefont.setColor(IndexedColors.GREEN.getIndex());

		// false -> red
		Font falsefont = workbook.createFont();
		falsefont.setColor(IndexedColors.RED.getIndex());
		Row row;
		int columnCount;
		row = sheet.createRow(rowCount);
		columnCount = 0;
		for (String o : data) {
			Cell cell = row.createCell(columnCount++);
			if (o.toString().equalsIgnoreCase("true")) {
				cell.setCellStyle(t_style);
				t_style.setFont(truefont);
			} else if (o.toString().equalsIgnoreCase("false")) {
				cell.setCellStyle(f_style);
				f_style.setFont(falsefont);
			}
			cell.setCellValue(o.toString());
		}
		return ++rowCount;
	}

	private static int writeHeader(List<String> header, int rowCount) {
		CellStyle h_style = workbook.createCellStyle();
		Font h_font = workbook.createFont();
		h_font.setFontHeightInPoints((short) 10);
		h_font.setBold(true);
		h_font.setColor(IndexedColors.BROWN.getIndex());
		
		Row row = sheet.createRow(rowCount);
		int columnCount = 0;
		for (String headerword : header) {
			Cell cell = row.createCell(columnCount++);
			cell.setCellStyle(h_style);
			h_style.setFont(h_font);
			cell.setCellValue(headerword);
		}
		return ++rowCount;
	}

	public static void writeAllData(String ExcelFileName) throws FileNotFoundException, IOException {
		FileOutputStream fos = new FileOutputStream(new File(ExcelFileName + ".xlsx"));
		workbook.write(fos);
		fos.close();
		System.out.println("Done");
	}
}
