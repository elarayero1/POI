package poiexampleEXCEL;

import java.io.FileNotFoundException;  
import java.io.FileOutputStream;  
import java.io.IOException;  
import java.io.OutputStream;  
import org.apache.poi.hssf.usermodel.HSSFCellStyle;  
import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.CellStyle;  
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.ss.usermodel.Sheet;  
import org.apache.poi.ss.usermodel.Workbook;  
  
public class AlignExample {  
    public static void main(String[] args) throws FileNotFoundException, IOException {  
        Workbook wb = new HSSFWorkbook();  
        try   {  
        	OutputStream os = new FileOutputStream("JavatpointAlign.xls");
            Sheet sheet = wb.createSheet("New Sheet");  
            Row row = sheet.createRow(0);  
            Cell cell = row.createCell(0);  
            cell.setCellValue("Javatpoint");  
            CellStyle cellStyle = wb.createCellStyle();  
            row = sheet.createRow(5);   
            cell = row.createCell(0);  
            // Top Left alignment   
            HSSFCellStyle style1 = (HSSFCellStyle) wb.createCellStyle();  
            sheet.setColumnWidth(0, 8000);  
            cell.setCellValue("Top Left");  
            cell.setCellStyle(style1);  
            row = sheet.createRow(6);   
            cell = row.createCell(1);  
            row.setHeight((short) 800);  
            wb.write(os);  
            System.out.println("excel creado exitosamente!");  
        }catch(Exception e) {  
            System.out.println(e.getMessage());  
        }  
    }  
  
}  