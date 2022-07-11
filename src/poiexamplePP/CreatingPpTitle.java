package poiexamplePP;

import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class CreatingPpTitle {
	public static void main(String[] args) {  
        XMLSlideShow ppt = new XMLSlideShow();  
        try (OutputStream os = new FileOutputStream("PowerPoint/JavatpointTitle.pptx")) {  
            XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);  
            XSLFSlideLayout titleLayout = defaultMaster.getLayout(SlideLayout.TITLE);  
            XSLFSlide slide = ppt.createSlide(titleLayout);  
            XSLFTextShape title = slide.getPlaceholder(0);  
            title.setText("David ");  
            ppt.write(os);  
            System.out.println("listo");
        }catch(Exception e) {  
            System.out.println(e);  
        }  
    }
}