package task13;

import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


	public class ReadEx {
	
		public static void main(String[] args) throws IOException {
													// open the book 
		XSSFWorkbook book = new XSSFWorkbook("C:\\Users\\HP\\eclipse-workspace\\Task13\\Data\\Book1.xlsx");

		
		
		XSSFSheet sheet = book.getSheetAt(0); // go to the sheet
		
		
		
		int rowCount = sheet.getLastRowNum();	// to get no of rows
		
		
		
		int columnCount = sheet.getRow(0).getLastCellNum(); // to get no of columns
		
		
		
		String[][] data = new String[rowCount][columnCount];  // array to store the value in row and column
		
		for(int i=1;i<=rowCount;i++) {  
			
			XSSFRow row = sheet.getRow(i);
			
			
			
			for(int j=0;j<columnCount;j++) {
				
				XSSFCell cell = row.getCell(j);
				
				
				
			System.out.println(cell.getStringCellValue()); // get the value
				
				
				
				data[i-1][j] = cell.getStringCellValue();  // store the value to array
				
			}
			System.out.println();
		}
		
		
		book.close();
	}
	}



	

