package poiexampleEXCEL;

import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderExtent;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;

public class DrawingBorderExample {
	public static void main(String[] args) {
		try (OutputStream os = new FileOutputStream("JavatpointDrawingBorder.xls")) {
			PropertyTemplate pt = new PropertyTemplate();
			pt.drawBorders(new CellRangeAddress(1, 2, 1, 2), BorderStyle.MEDIUM, BorderExtent.ALL);
			pt.drawBorders(new CellRangeAddress(5, 6, 1, 2), BorderStyle.MEDIUM, BorderExtent.OUTSIDE);
			pt.drawBorders(new CellRangeAddress(5, 6, 1, 2), BorderStyle.THIN, BorderExtent.INSIDE);
			pt.drawBorders(new CellRangeAddress(9, 10, 1, 3), BorderStyle.MEDIUM, IndexedColors.GREEN.getIndex(),
					BorderExtent.OUTSIDE);
			pt.drawBorders(new CellRangeAddress(9, 10, 1, 3), BorderStyle.MEDIUM, IndexedColors.BLUE.getIndex(),
					BorderExtent.INSIDE_VERTICAL);
			pt.drawBorders(new CellRangeAddress(9, 10, 1, 3), BorderStyle.MEDIUM, IndexedColors.RED.getIndex(),
					BorderExtent.INSIDE_HORIZONTAL);
			pt.drawBorders(new CellRangeAddress(10, 10, 2, 2), BorderStyle.NONE, BorderExtent.ALL);
			Workbook wb = new HSSFWorkbook();
			Sheet sheet = wb.createSheet("Sheet");
			pt.applyBorders(sheet);
			wb.write(os);
			System.out.println("creado!");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}