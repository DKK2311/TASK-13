package writeoperation;

import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class readoperation {

	public static void main(String[] args) throws IOException {

		 XSSFWorkbook book =new XSSFWorkbook("D:\\dharaneesh\\FINAL YEAR PROJECT\\guvi\\Task-13\\vk.xlsx");
	     XSSFSheet sheet = book.getSheetAt(0);    
	     
	     int rowCount = sheet.getLastRowNum();                
	     int columnCount =sheet.getRow(0).getLastCellNum();     
	     
	   //creating an array
	     
	     Object [][]  data =new Object[rowCount][columnCount];      
	     
	   //getting the row
	     
	     for(int i=0;i<rowCount;i++) {              
	    	 XSSFRow row= sheet.getRow(i);
	    	 
	    //getting the cell
	    	 
	    	 for(int j=0;j<columnCount;j++) 
	    	 { 
	    		 
	    		XSSFCell cell = row.getCell(j);
	    		
	    		//getting the cell value and putting into a array
	    		
	    		data[i][j] = cell.getStringCellValue();         
	    		
	    		 //printing the value
	    		System.out.println(cell.getStringCellValue());  
	    		 
	    	 }
	     }
	     
	   //closing the book
	       book.close();                         
		
	}

}
