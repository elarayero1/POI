package poiexamplePP;

import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.poi.xslf.usermodel.XMLSlideShow;

public class CreatingPptExample {
	public static void main(String[] args) {
		XMLSlideShow ppt = new XMLSlideShow();
		try (OutputStream os = new FileOutputStream("PowerPoint/Javatpoint.pptx")) {
			ppt.write(os);
			System.out.println("listo");
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}