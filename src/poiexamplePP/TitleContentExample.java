package poiexamplePP;

import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class TitleContentExample {
	public static void main(String[] args) {
		XMLSlideShow ppt = new XMLSlideShow();
		try (OutputStream os = new FileOutputStream("PowerPoint/JavatpointTitle2.pptx")) {
			XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);
			XSLFSlideLayout tc = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
			XSLFSlide slide = ppt.createSlide(tc);
			XSLFTextShape title = slide.getPlaceholder(0);
			title.setText("Title here David");
			XSLFTextShape body = slide.getPlaceholder(1);
			body.clearText();
			body.addNewTextParagraph().addNewTextRun().setText("This is a new slide created using Java program.");
			ppt.write(os);
			System.out.println("ok");
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}
