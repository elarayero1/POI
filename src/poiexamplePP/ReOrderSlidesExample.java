package poiexamplePP;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class ReOrderSlidesExample {
	public static void main(String args[]) throws IOException {
		try (XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("PowerPoint/Javatpointpaginas.pptx"))) {
			// Getting all the slides
			List<XSLFSlide> slides = ppt.getSlides();
			// Selecting the second slide
			XSLFSlide secondslide = slides.get(1);
			// Getting on the top
			ppt.setSlideOrder(secondslide, 0);
			// Writing Modifications
			FileOutputStream out = new FileOutputStream("PowerPoint/JavatpointCambiaDiapositiva.pptx");
			ppt.write(out);
			System.out.println("ok");
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}