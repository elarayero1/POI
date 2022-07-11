package poiexampleEXCEL;

import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ShiftRowsExample {
	public static void main(String[] args) {
		try (OutputStream os = new FileOutputStream("JavatpointShiftRows.xls")) {
			Workbook wb = new HSSFWorkbook();
			Sheet sheet = wb.createSheet("Sheet");
			Row row = sheet.createRow(1);
			Row row2 = sheet.createRow(2);
			Cell cell = row.createCell(1);
			Cell cell2 = row2.createCell(1);
			cell2.setCellValue("Hello, Javatpoint");
			sheet.shiftRows(1, 2, -1);
			wb.write(os);
			System.out.println("creado");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}
