package poiexampleEXCEL;

import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class CellComments {
	public static void main(String[] args) throws IOException {
		try (FileOutputStream out = new FileOutputStream("JavatpointComent.xls")) {
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet sheet = wb.createSheet("Comment Sheet");
			HSSFPatriarch hpt = sheet.createDrawingPatriarch();
			HSSFCell cell1 = sheet.createRow(3).createCell(1);
			cell1.setCellValue("Excel Comment Example");
			// Setting size and position of the comment in worksheet
			HSSFComment comment1 = hpt.createComment(new HSSFClientAnchor(0, 0, 0, 0, (short) 4, 2, (short) 6, 5));
			// Setting comment text
			comment1.setString(new HSSFRichTextString("It is a comment"));
			// Associating comment to the cell
			cell1.setCellComment(comment1);
			wb.write(out);
			System.out.println("creado");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}
