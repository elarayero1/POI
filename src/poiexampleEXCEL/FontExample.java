package poiexampleEXCEL;

import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class FontExample {
	public static void main(String[] args) {
		try (OutputStream fileOut = new FileOutputStream("JavatpointFont.xls")) {
			Workbook wb = new HSSFWorkbook(); // Creating a workbook
			Sheet sheet = wb.createSheet("Sheet"); // Creating a sheet
			Row row = sheet.createRow(1); // Creating a row
			Cell cell = row.createCell(1); // Creating a cell
			CellStyle style = wb.createCellStyle(); // Creating Style
			cell.setCellValue("Hello, Javatpoint!");
			// Creating Font and settings
			Font font = wb.createFont();
			font.setFontHeightInPoints((short) 11);
			font.setFontName("Courier New");
			font.setItalic(true);
			font.setStrikeout(true);
			// Applying font to the style
			style.setFont(font);
			// Applying style to the cell
			cell.setCellStyle(style);
			wb.write(fileOut);
			System.out.println("excel creado!");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}
