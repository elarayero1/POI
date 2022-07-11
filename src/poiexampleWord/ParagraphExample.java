package poiexampleWord;

import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ParagraphExample {
	public static void main(String[] args) {
		XWPFDocument doc = new XWPFDocument();
		try (OutputStream os = new FileOutputStream("Word/JavatpointParagraph.doc")) {
			XWPFParagraph paragraph = doc.createParagraph();
			XWPFRun run = paragraph.createRun();
			run.setText("Hello, This is javatpoint. This paragraph is written " + "by using XWPFParagrah.");
			doc.write(os);
			System.out.println("creado");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}
