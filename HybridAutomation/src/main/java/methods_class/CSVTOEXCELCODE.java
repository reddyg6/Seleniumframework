package methods_class;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CSVTOEXCELCODE {

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
