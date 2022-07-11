package poiexampleEXCEL;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class NewCellExample {

	 public static void main(String[] args) throws FileNotFoundException, IOException {  
	        Workbook wb = new HSSFWorkbook();  
	        try  (OutputStream os = new FileOutputStream("NewCellExample.xls")) {  
	            Sheet sheet = wb.createSheet("New Sheet");  
	            Row row     = sheet.createRow(2);  
	            Cell cell   = row.createCell(5);  
	            cell.setCellValue("Javatpoint");  
	            wb.write(os);  
	        }catch(Exception e) {  
	            System.out.println(e.getMessage());  
	        }  
	    }  
}