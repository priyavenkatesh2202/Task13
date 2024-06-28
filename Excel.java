package task13;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {

	
		public static void main(String[] args) throws IOException {
			// TODO Auto-generated method stub
			
			
			
			XSSFWorkbook book = new XSSFWorkbook();
			
			
			
			XSSFSheet sheet = book.createSheet("SHEET"); // creation of sheet
			
			
			Object[][] data = {
					{"Name","Age","Email"},
					{"John Doe","30","john@test.com"},
					{"Jane Doe","28","john@test.com"},	 		// data to be entered in the cell
					{"Bob Smith","35","jacky@example.com"},
					{"Swapnil","37","swapnil@example.com"},
					
					 };
			
			
			int rowcount= 0 ;
			for(Object[] row1 : data) {
				XSSFRow row = sheet.createRow(rowcount++);	// for data to be entered in row
				

				int columnCount=0;

				

				for(Object col:row1) {

					XSSFCell cell = row.createCell(columnCount++); // for data to be entered in column

					
					if(col instanceof String) {				
						cell.setCellValue((String)col);			// for string input
					}else if (col instanceof Integer) {
						cell.setCellValue((Integer) col);		// for int age input
					}
				}

			}

			try {
				 FileOutputStream output = new FileOutputStream ("C:\\Users\\HP\\eclipse-workspace\\Task13\\src\\main\\java\\task13\\Sheet1.xlsx");
				
				
				
				book.write(output);
			} catch (Exception e) 
			{
				e.printStackTrace();
			}
	book.close();
			
		}	}



			
			
			
	
	


