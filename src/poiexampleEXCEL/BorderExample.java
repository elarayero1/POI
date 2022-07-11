package poiexampleEXCEL;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class BorderExample {

	public static void main(String[] args) throws FileNotFoundException, IOException {  
        Workbook wb = new HSSFWorkbook();  
        Sheet sheet = wb.createSheet("Sheet");  
        Row row     = sheet.createRow(1);  
        Cell cell   = row.createCell(1);  
        cell.setCellValue(101);  
        // Styling border of cell.  
        CellStyle style = wb.createCellStyle();  
        style.setBorderBottom(BorderStyle.THIN);  
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());  
        style.setBorderRight(BorderStyle.THIN);  
        style.setRightBorderColor(IndexedColors.BLUE.getIndex());  
        style.setBorderTop(BorderStyle.MEDIUM_DASHED);  
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());  
        cell.setCellStyle(style);  
        try (OutputStream fileOut = new FileOutputStream("JavatpointBorder.xls")) {  
            wb.write(fileOut);  
            System.out.println("excel creado exitosamente!");
        }catch(Exception e) {  
            System.out.println(e.getMessage());  
        }  
}
	
}
