package poiexampleEXCEL;

import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class NewlineExample {
	public static void main(String[] args) {
		try (OutputStream fileOut = new FileOutputStream("JavatpointNewline.xls")) {
			Workbook wb = new HSSFWorkbook();
			Sheet sheet = wb.createSheet("Sheet");
			Row row = sheet.createRow(1);
			Cell cell = row.createCell(1);
			cell.setCellValue("This is first line and \n this is second line");
			CellStyle cs = wb.createCellStyle();
			cs.setWrapText(true);
			cell.setCellStyle(cs);
			row.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
			sheet.autoSizeColumn(2);
			wb.write(fileOut);
			System.out.println("creado");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}
