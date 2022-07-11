package poiexampleEXCEL;

import java.io.FileInputStream;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class GettingCellContentExample {
	public static void main(String[] args) {
		try (InputStream inp = new FileInputStream("Javatpoint.xls")) {
			Workbook wb = WorkbookFactory.create(inp);
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.getRow(2);
			Cell cell = row.getCell(3);
			if (cell != null)
				System.out.println("Data: " + cell);
			else
				System.out.println("Cell is empty");
			//wb.write(os);
			System.out.println("listo");
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}