package poiexampleEXCEL;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;

public class HyperlinkFormula {
	public static void main(String[] args) throws IOException {
		try (HSSFWorkbook wb = new HSSFWorkbook()) {
			HSSFSheet sheet = wb.createSheet("new sheet");
			HSSFRow row = sheet.createRow(0);
			HSSFCell cell = row.createCell(0);
			cell.setCellType(CellType.FORMULA);
			cell.setCellFormula("HYPERLINK(\"http://https://www.javatpoint.com/apache-poi-tutorial\", \"click here\")");
			try (FileOutputStream fileOut = new FileOutputStream("JavatpointLink.xls")) {
				wb.write(fileOut);
				System.out.println("creado");
			}
		}
	}
}
