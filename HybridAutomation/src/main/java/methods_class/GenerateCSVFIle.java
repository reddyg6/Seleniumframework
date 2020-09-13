package methods_class;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
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

import com.opencsv.CSVWriter;

import io.github.bonigarcia.wdm.WebDriverManager;

public class GenerateCSVFIle {
	static WebDriver driver;
	static String Data;
	static int row;

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		System.out.print("Current Time in milliseconds = ");
		long startTime = System.nanoTime();

		WebDriverManager.chromedriver().setup();
		driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
		driver.manage().window().maximize();
		driver.get("https://www.barclays.co.uk/insurance/home-insurance/");

		String CSV_File_Path = "D:/UsersuserDesktopq1.csv";
		FileWriter outputfile = new FileWriter(new File(CSV_File_Path));
		CSVWriter writer = new CSVWriter(outputfile);
		List<String[]> data = new ArrayList<String[]>();
		List<WebElement> H2 = driver.findElements(By.tagName("h2"));

		data.add(new String[] { "H2_Color", "H2_Color_status", "H2_Font_Family", "Details" });

		for (int i = 1; i < H2.size(); i++) { // int i = row - 1; if
			if (H2.get(i).isDisplayed() == true) {

				/*
				 * Actions ac = new Actions(driver);
				 * ac.moveToElement(H2.get(i)).build().perform();
				 */

				String h2_color_value = H2.get(i).getCssValue("color");
				String hexcolor = Color.fromString(h2_color_value).asHex();
				boolean color = hexcolor.equalsIgnoreCase("#00395D");

				String h2_Fontfamily = H2.get(i).getCssValue("font-family").replace(",", "");
				boolean fs = h2_Fontfamily
						.equalsIgnoreCase("\"Expert Sans Light\", \"Trebuchet MS\", Arial, Verdana, sans-serif");
				data.add(new String[] { hexcolor, "" + color + "", h2_Fontfamily, "" + fs + "" });

			}
		}

		List<WebElement> H3 = driver.findElements(By.tagName("h3"));

		data.add(new String[] { "H3_Color", "H3_Color_status", "H3_Font_Family" });

		for (int i = 1; i < H3.size(); i++) { // int i = row - 1; if
			if (H3.get(i).isDisplayed() == true) {

				/*
				 * Actions ac = new Actions(driver);
				 * ac.moveToElement(H2.get(i)).build().perform();
				 */

				String h3_color_value = H3.get(i).getCssValue("color");
				String hexcolor = Color.fromString(h3_color_value).asHex();
				boolean color = hexcolor.equalsIgnoreCase("#00395D");

				String h3_Fontfamily = H3.get(i).getCssValue("font-family");

				data.add(new String[] { hexcolor, "" + color + "", h3_Fontfamily });
			}
		}

		writer.writeAll(data);
		writer.close();
		long endTime = System.nanoTime();
		long timeElapsed = endTime - startTime;

		System.out.print("Current Time in milliseconds = ");
		System.out.println(timeElapsed / 1000000);

		driver.quit();

		csvtoExcel_star("D:/UsersuserDesktopq1.csv", "D:\\Javatpoint.xlsx");
		System.out.println("Successfully file converted into excel");

	}

	private static void csvtoExcel_star(String CSVFilePath, String ExcelFilePath)
			throws FileNotFoundException, IOException {

		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("sheet1");
		// true -> green
		CellStyle t_style = workbook.createCellStyle();
		CellStyle f_style = workbook.createCellStyle();
		Font truefont = workbook.createFont();
		truefont.setColor(IndexedColors.GREEN.getIndex());

		// false -> red
		Font falsefont = workbook.createFont();
		falsefont.setColor(IndexedColors.RED.getIndex());

		BufferedReader br = new BufferedReader(new FileReader(new File(CSVFilePath)));
		String line = br.readLine();
		int rowcount = 0;
		while (line != null) {
			Row row = sheet.createRow(rowcount);
			int cellcount = 0;
			for (String word : line.split(",", -1)) {
				Cell cell = row.createCell(cellcount);
				word.toString();
				if (word.startsWith("=")) {
					cell.setCellType(Cell.CELL_TYPE_STRING);
					word = word.replaceAll("\"", "");
					word = word.replaceAll("=", "");

					if (word.equalsIgnoreCase("true")) {

						cell.setCellStyle(t_style);
						t_style.setFont(truefont);
						cell.setCellValue(word);
					} else if (word.equalsIgnoreCase("false")) {

						cell.setCellStyle(f_style);
						f_style.setFont(falsefont);
						cell.setCellValue(word);
					} else {
						cell.setCellValue(word);
					}

				} else if (word.startsWith("\"")) {
					word = word.replaceAll("\"", "");
					cell.setCellType(Cell.CELL_TYPE_STRING);
					cell.setCellValue(word);
					if (word.equalsIgnoreCase("true")) {

						cell.setCellStyle(t_style);
						t_style.setFont(truefont);
						cell.setCellValue(word);
					} else if (word.equalsIgnoreCase("false")) {
						cell.setCellStyle(f_style);
						f_style.setFont(falsefont);
						cell.setCellValue(word);
					} else {
						cell.setCellValue(word);
					}
				} else {
					word = word.replaceAll("\"", "");
					cell.setCellType(Cell.CELL_TYPE_NUMERIC);
					cell.setCellValue(word);

				}

				cellcount++;
			}
			rowcount++;
			line = br.readLine();
		}
		FileOutputStream fos = new FileOutputStream(new File(ExcelFilePath));
		workbook.write(fos);
		fos.close();
		System.out.println("Done");
	}

}
