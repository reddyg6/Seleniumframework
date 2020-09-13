package methods_class;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelApiTest {

	
	
	public FileInputStream fis = null;
    public FileOutputStream fos = null;
    public XSSFWorkbook workbook = null;
    public XSSFSheet sheet = null;
    public XSSFRow row = null;
    public XSSFCell cell = null;
    String xlFilePath;

    public ExcelApiTest(String xlFilePath) throws Exception
    {
        this.xlFilePath = xlFilePath;
        fis = new FileInputStream(xlFilePath);
        workbook = new XSSFWorkbook(fis);
        fis.close();
    }
    
    //Data = "\"" + h3_value + "\",\"" + hexabgc + "\",\"" + color + "\",\"" + FS + "\"\r\n";
	/*
	 * String Filepath = "D:\\Testdata.csv"; Data =
	 * "\"h3heading\",\"h3_backgroungcolor\",\"h3_color\",\"h3_Fontsize\"\n";
	 * ExcelApiTest.writedataintocsv(Data, Filepath);
	 * ExcelApiTest.writedataintocsv(Data, Filepath);
	 */

    public boolean setCellData(String sheetName, String colName, int rowNum, String value)
    {
        try
        {
            int col_Num = -1;
            sheet = workbook.getSheet(sheetName);

            row = sheet.getRow(0);
            for (int i = 0; i < row.getLastCellNum(); i++) {
                if (row.getCell(i).getStringCellValue().trim().equals(colName))
                {
                    col_Num = i;
                }
            }

            sheet.autoSizeColumn(col_Num);
            row = sheet.getRow(rowNum - 1);
            if(row==null)
                row = sheet.createRow(rowNum - 1);

            cell = row.getCell(col_Num);
            if(cell == null)
                cell = row.createCell(col_Num);

            cell.setCellValue(value);

            fos = new FileOutputStream(xlFilePath);
            workbook.write(fos);
            fos.close();
        }
        catch (Exception ex)
        {
            ex.printStackTrace();
            return  false;
        }
        return true;
    }
    
	// returns the data from a cell
	public String getCellData(String sheetName,String colName,int rowNum){
		try{
			if(rowNum <=0)
				return "";
		
		int index = workbook.getSheetIndex(sheetName);
		int col_Num=-1;
		if(index==-1)
			return "";
		
		sheet = workbook.getSheetAt(index);
		row=sheet.getRow(0);
		for(int i=0;i<row.getLastCellNum();i++){
			//System.out.println(row.getCell(i).getStringCellValue().trim());
			if(row.getCell(i).getStringCellValue().trim().equals(colName.trim())){
				col_Num=i;
				break;
			}
		}
		if(col_Num==-1)
			return "";
		
		sheet = workbook.getSheetAt(index);
		row = sheet.getRow(rowNum);
		if(row==null)
			return "";
		cell = row.getCell(col_Num);
		
		if(cell==null)
			return "";
		//System.out.println(cell.getCellType());
		if(cell.getCellType()==Cell.CELL_TYPE_STRING){
			
			
			 cell.setCellType(Cell.CELL_TYPE_STRING);
			 return cell.getStringCellValue();
			
		}
			 
//		else if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC){
//			String celltText = cell.getNumericCellValue()
//		}
		else if( cell.getCellType()==Cell.CELL_TYPE_FORMULA || cell.getCellType()==Cell.CELL_TYPE_NUMERIC  ){
		
			 // System.out.println("Data type"+cell.getNumericCellValue());
			
			  String cellText  = String.valueOf(cell.getRawValue());
			  
			  //String cellText  = String.valueOf(cell.getStringCellValue());
			  if (HSSFDateUtil.isCellDateFormatted(cell)) {
		           // format in form of M/D/YY
				  double d = cell.getNumericCellValue();

				  Calendar cal =Calendar.getInstance();
				  cal.setTime(HSSFDateUtil.getJavaDate(d));
		          //cellText =(String.valueOf(cal.get(Calendar.YEAR))).substring(2);
				  int mon = cal.get(Calendar.MONTH);
				  mon = mon+1;
		          cellText = cal.get(Calendar.DAY_OF_MONTH) + "/" + mon + "/" + cal.get(Calendar.YEAR);

		         }

		return cellText;
		}else if(cell.getCellType()==Cell.CELL_TYPE_BLANK)
		      return ""; 
		  else 
			  return String.valueOf(cell.getBooleanCellValue());
		
		}
		catch(Exception e){
			
			e.printStackTrace();
			return "row "+rowNum+" or column "+colName +" does not exist in xls";
		}
	}
    
	 public String readCellData(String sheetName, String colName, int rowNum)
	    {
	        try
	        {
	            int col_Num = -1;
	            sheet = workbook.getSheet(sheetName);
	            row = sheet.getRow(0);
	            for(int i = 0; i < row.getLastCellNum(); i++)
	            {
	                if(row.getCell(i).getStringCellValue().trim().equals(colName.trim()))
	                    col_Num = i;
	            }
	 
	            row = sheet.getRow(rowNum - 1);
	            cell = row.getCell(col_Num);
	 
	            if(cell.getCellTypeEnum() == CellType.STRING)
	                return cell.getStringCellValue();
	            else if(cell.getCellTypeEnum() == CellType.NUMERIC || cell.getCellTypeEnum() == CellType.FORMULA)
	            {
	                String cellValue = String.valueOf(cell.getNumericCellValue());
	                if(HSSFDateUtil.isCellDateFormatted(cell))
	                {
	                    DateFormat df = new SimpleDateFormat("dd/MM/yy");
	                    Date date = cell.getDateCellValue();
	                    cellValue = df.format(date);
	                }
	                return cellValue;
	            }else if(cell.getCellTypeEnum() == CellType.BLANK)
	                return "";
	            else
	                return String.valueOf(cell.getBooleanCellValue());
	        }
	        catch(Exception e)
	        {
	            e.printStackTrace();
	            return "row "+rowNum+" or column "+colName+" does not exist  in Excel";
	        }
	    }
    
    
    public static void writedataintocsv(String Data, String Filepath) throws IOException {
    
    	FileUtils.writeStringToFile(new File(Filepath), Data,true);
    }
      
    public static void writedataintoexcel() throws Exception{
    
    	String NameOfSheetName="";
    	SXSSFWorkbook workbook= new SXSSFWorkbook(100);  
    	SXSSFSheet sheetName=workbook.createSheet(NameOfSheetName);
    	for (int r = 0; r < 55555; r++) {  
            Row row = sheetName.createRow(r);  
   
            // Create a Font for styling header cells
            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setFontHeightInPoints((short) 16);
            headerFont.setColor(IndexedColors.RED.getIndex());
            
            // Create a CellStyle with the font
            CellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFont(headerFont);

            // Create a Row
            Row headerRow = sheetName.createRow(0);
            
//iterating c number of columns  
            for (int c = 0; c < 5; c++) {  
                Cell cell = row.createCell(c);  
                cell.setCellValue("Cell " + r + " " + c);  
            }  
            if ( r % 1000 == 0) {
                System.out.println(r);
            }
        }  
   
        FileOutputStream out = new FileOutputStream("");
        workbook.write(out);
        out.flush();
        out.close();
    } 
    	

    
    
}
