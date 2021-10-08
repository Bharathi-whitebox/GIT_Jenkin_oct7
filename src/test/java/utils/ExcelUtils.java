package utils;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	
	static String projectPath;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	
	public ExcelUtils(String excelPath , String sheetName) {
	
		try { 
			
			 projectPath = System.getProperty("user.dir");
			 workbook     = new XSSFWorkbook(excelPath);
			 sheet        = workbook.getSheet(sheetName);
			 
		}catch ( Exception e){
			
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();			
			
		}
		
	}
	
	//public static void main(String [] args) {
		
//		getRowCount() ;
//		getCellDataString(1,0); 
//		getCellDataNumber(1,1);
//		
//	}
	
	public static int getRowCount() {
		 int rowCount = 0;
		
		try { 
			 rowCount = sheet.getPhysicalNumberOfRows();
		     System.out.println("No of Rows: " + rowCount);
				
		}catch ( Exception e){
			
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();			
			
		}
		return rowCount;
		
	}
	
	public static int getColCount() {
		
		int colCount = 0;
		try { 
			  colCount = sheet.getRow(0).getPhysicalNumberOfCells();
		      System.out.println("No of Colunm: " + colCount);
				
		}catch ( Exception e){
			
			System.out.println(e.getMessage());
			System.out.println(e.getCause());
			e.printStackTrace();			
			
		}
		return colCount;
		
	}
	
	public static String getCellDataString(int rowNum , int colNum)  {
		
			String cellData = null;
		
			try { 
				  cellData =  sheet.getRow(rowNum).getCell(colNum).getStringCellValue();   // cell 00 
		//		  System.out.println(cellData);
			 
			}catch ( Exception e){
		
				System.out.println(e.getMessage());
				System.out.println(e.getCause());
				e.printStackTrace();		
	        }
			return cellData;
	}
	
	public static void getCellDataNumber(int rowNum , int colNum)  {
		
			try { 		 
				 Double cellDataNum =  sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
				// System.out.println(cellDataNum);
			 
			}catch ( Exception e){
		
				System.out.println(e.getMessage());
				System.out.println(e.getCause());
				e.printStackTrace();		
	    }
	}
}
