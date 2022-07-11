package poiexampleEXCEL;

import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ShapeExample {
	public static void main(String[] args) {
		Workbook wb = new HSSFWorkbook();
		try (OutputStream os = new FileOutputStream("JavatpointShape.xls")) {
			Sheet sheet = wb.createSheet("Sheet");
			Row row = sheet.createRow(4); // Creating a row
			Cell cell = row.createCell(1); // Creating a cell
			HSSFPatriarch patriarch = (HSSFPatriarch) sheet.createDrawingPatriarch();
			HSSFClientAnchor a = new HSSFClientAnchor(0, 0, 1023, 255, (short) 1, 0, (short) 1, 0);
			HSSFSimpleShape shape = patriarch.createSimpleShape(a);
			shape.setShapeType(HSSFSimpleShape.OBJECT_TYPE_OVAL);
			wb.write(os);
			System.out.println("creado");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}