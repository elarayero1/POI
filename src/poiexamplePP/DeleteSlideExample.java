package poiexamplePP;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.xslf.usermodel.XMLSlideShow;

public class DeleteSlideExample {
	public static void main(String[] args) {
		try (XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("PowerPoint/Javatpointpaginas.pptx"))) {
			ppt.removeSlide(0);
			FileOutputStream out = new FileOutputStream("PowerPoint/JavatpointpaginasDelete.pptx");
			ppt.write(out);
			System.out.println("ok");
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}
