package poiexampleEXCEL;


import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
public class ImageExample {

	public static void main(String[] args) {
		try (OutputStream os = new FileOutputStream("JavatpointImg.xls")) {
			Workbook wb = new HSSFWorkbook();
			
			//final FileInputStream stream =  new FileInputStream( "C:/Users/jdvelasquez/Pictures/1.PNG" );
			
			String imagePath = "C:/Users/jdvelasquez/Pictures/1.PNG";
			
			Sheet sheet = wb.createSheet("Sheet");
			Row row = sheet.createRow(2);
			Drawing drawing = sheet.createDrawingPatriarch();
			  ClientAnchor clientAnchor = wb.getCreationHelper().createClientAnchor();
			  clientAnchor.setRow1(2);
			  clientAnchor.setCol1(4);
			  Picture picture = drawing.createPicture(clientAnchor,2);
			  picture.resize();
			
			wb.write(os);
			System.out.println("creado");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}