package poiexamplePP;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

public class ReOrderSlidesExampleImage {
	public static void main(String[] args) throws FileNotFoundException, IOException {
		XMLSlideShow ppt = new XMLSlideShow();
		try (OutputStream os = new FileOutputStream("PowerPoint/JavatpointImg.pptx")) {
			XSLFSlide slide = ppt.createSlide();
			byte[] pictureData = IOUtils.toByteArray(new FileInputStream("C:/Users/jdvelasquez/Pictures/1.PNG"));
			XSLFPictureData pd = ppt.addPicture(pictureData, XSLFPictureData.PictureType.PNG);
			XSLFPictureShape pic = slide.createPicture(pd);
			ppt.write(os);
			System.out.println("ok");
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}