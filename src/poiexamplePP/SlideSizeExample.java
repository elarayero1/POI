package poiexamplePP;

import java.io.FileInputStream;
import org.apache.poi.xslf.usermodel.XMLSlideShow;

public class SlideSizeExample {
	public static void main(String[] args) {
		try (XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("PowerPoint/Javatpointpaginas.pptx"))) {
			java.awt.Dimension pgsize = ppt.getPageSize();
			int width = pgsize.width; // slide width in points
			int height = pgsize.height; // slide height in points
			System.out.println("width: " + width);
			System.out.println("height: " + height);
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}