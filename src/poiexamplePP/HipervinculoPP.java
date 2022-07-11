package poiexamplePP;

import java.io.FileNotFoundException;  
import java.io.FileOutputStream;  
import java.io.IOException;  
import java.io.OutputStream;  
import org.apache.poi.xslf.usermodel.SlideLayout;  
import org.apache.poi.xslf.usermodel.XMLSlideShow;  
import org.apache.poi.xslf.usermodel.XSLFHyperlink;  
import org.apache.poi.xslf.usermodel.XSLFSlide;  
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;  
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;  
import org.apache.poi.xslf.usermodel.XSLFTextRun;  
import org.apache.poi.xslf.usermodel.XSLFTextShape;  
public class HipervinculoPP {  
    public static void main(String[] args) throws FileNotFoundException, IOException {  
        XMLSlideShow ppt = new XMLSlideShow();  
        try (OutputStream os = new FileOutputStream("PowerPoint/JavatpointLink.pptx")) {         
            // Setting layout  
            XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);  
            XSLFSlideLayout tc = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);  
            XSLFSlide slide = ppt.createSlide(tc);  
            // Setting title  
            XSLFTextShape title = slide.getPlaceholder(0);  
            title.setText("Hyperlink Example");  
            // Setting body  
            XSLFTextShape body = slide.getPlaceholder(1);  
            body.clearText();  
            XSLFTextRun r = body.addNewTextParagraph().addNewTextRun();  
            r.setText("Click here to visit Javatpoint.");  
            XSLFHyperlink link = r.createHyperlink();  
            link.setAddress("https://www.javatpoint.com");  
            ppt.write(os); 
            System.out.println("OK");
        }catch(Exception e) {  
             System.out.println(e);  
         }  
    }  
} 
