package poiexampleEXCEL;

import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class MergingCellExample {
	public static void main(String[] args) {
		try (OutputStream fileOut = new FileOutputStream("JavatpointMergin.xls")) {
			Workbook wb = new HSSFWorkbook();
			Sheet sheet = wb.createSheet("Sheet");
			Row row = sheet.createRow(1);
			Cell cell = row.createCell(1);
			cell.setCellValue("Two cells have merged");
			// Merging cells by providing cell index
			sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 2));
			wb.write(fileOut);
			System.out.println("excel creado exitosamente!");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}
