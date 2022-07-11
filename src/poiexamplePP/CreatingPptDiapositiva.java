package poiexamplePP;

import java.io.FileOutputStream;  
import java.io.OutputStream;  
import org.apache.poi.xslf.usermodel.XMLSlideShow;  
import org.apache.poi.xslf.usermodel.XSLFSlide; 

public class CreatingPptDiapositiva {
	public static void main(String[] args) {
		XMLSlideShow ppt = new XMLSlideShow();
		try (OutputStream os = new FileOutputStream("PowerPoint/JavatpointDiapositivaOne.pptx")) {
			XSLFSlide slide = ppt.createSlide();  
			ppt.write(os);
			System.out.println("listo");
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}
