package excelwrite;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Writeexcel {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		// TODO Auto-generated method stub
		XSSFWorkbook book=new XSSFWorkbook(); //creating a book
		XSSFSheet sheet = book.createSheet();	//creating a sheet

		Object[] [] data = {	 //creating a data

				{"Name","Age","City"},
				{"Dharaneesh","26","Erode"},
				{"Dhamotharan","66","Erode"},
				{"Billa","55","Erode"},
				
	};
		int rowCount=0; // initializing at row count

		for(Object[] row : data) { 

			XSSFRow createRow = sheet.createRow(rowCount++);

		int columnCount=0;  // initializing at Column count

		for(Object column: row) {                   

			XSSFCell cell = createRow.createCell(columnCount++);

			if(column instanceof String)
			{ 
				cell.setCellValue((String) column);
			}
			else if(column instanceof Integer)
			{
				cell.setCellValue((Integer) column);
			} 

			try(                                                 
					FileOutputStream output = new FileOutputStream("D:\\dharaneesh\\FINAL YEAR PROJECT\\guvi\\Task-13\\vk.xlsx");){
					book.write(output);           
				}

			}
		}

	}

}
