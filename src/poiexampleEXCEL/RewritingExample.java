package poiexampleEXCEL;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class RewritingExample {
	public static void main(String[] args)
			throws FileNotFoundException, IOException, EncryptedDocumentException, InvalidFormatException {
		try (InputStream inp = new FileInputStream("C:/Fabrica/bvc/Desarrollo/IntegralRRHH/POI/Javatpoint.xls")) {
			Workbook wb = WorkbookFactory.create(inp);
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(2);
			Cell cell = row.getCell(3);
			if (cell == null)
				cell = row.createCell(3);
			cell.setCellType(CellType.STRING);
			cell.setCellValue("101");
			try (OutputStream fileOut = new FileOutputStream("Javatpoint.xls")) {
				wb.write(fileOut);
			}
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}
