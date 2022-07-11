 package poiexampleWord;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.formula.functions.Replace;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ReadingText {
	public static void main(String[] args) {
		
		XWPFDocument doc = new XWPFDocument();
		
		try (FileInputStream fis = new FileInputStream("Word/JavatpointAlingnin.docx")) {
			XWPFDocument file = new XWPFDocument(OPCPackage.open(fis));
			XWPFWordExtractor ext = new XWPFWordExtractor(file);
			System.out.println(ext.getText());
			
			XWPFParagraph paragraph = doc.createParagraph();
			
			paragraph.setAlignment(ParagraphAlignment.DISTRIBUTE);
			XWPFRun run = paragraph.createRun();
			run.setText(ext.getText());
			
			String text = run.getText(0);
            if (text != null && text.contains("$david")) {
                text = text.replace("$david", "Jimenaun");
                run.setText(text, 0);
            }
			
			OutputStream os = new FileOutputStream("Word/JavatpointAlingninXXX.docx");
			
			doc.write(os);
			System.out.println("ok");
			
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}