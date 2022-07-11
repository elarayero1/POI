package poiexampleEXCEL;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Calendar;
import java.util.Date;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class CellTypeExample {
	public static void main(String[] args) {
		try (OutputStream os = new FileOutputStream("JavatpointCellType.xls")) {
			Workbook wb = new HSSFWorkbook();
			Sheet sheet = wb.createSheet("Sheet");
			Row row = sheet.createRow(2);
			row.createCell(0).setCellValue(1.1); // Float value
			row.createCell(1).setCellValue(" " + new Date()); // Date type
			row.createCell(2).setCellValue(" " + Calendar.getInstance());// Calendar
			row.createCell(3).setCellValue("a string value"); // String
			row.createCell(4).setCellValue(true); // Boolean
			row.createCell(5).setCellType(CellType.ERROR); // Error
			wb.write(os);
			System.out.println("creado");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}