package poiexamplePP;

import java.io.FileInputStream;
import org.apache.poi.xslf.usermodel.XMLSlideShow;

public class SlideSizeExample2 {

	public static void main(String[] args) {
		try (XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("PowerPoint/Javatpointpaginas.pptx"))) {
			java.awt.Dimension pgsize = ppt.getPageSize();
			int width = pgsize.width; // slide width in points
			int height = pgsize.height; // slide height in points
			System.out.println("width: " + width);
			System.out.println("height: " + height);
			ppt.setPageSize(new java.awt.Dimension(1024, 768));
			java.awt.Dimension newpgsize = ppt.getPageSize();
			System.out.println("\nSlide size after setting new size.");
			System.out.println("width: " + newpgsize.width);
			System.out.println("height: " + newpgsize.height);
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}