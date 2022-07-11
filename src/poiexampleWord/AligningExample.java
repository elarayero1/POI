package poiexampleWord;

import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class AligningExample {
	public static void main(String[] args) {
		XWPFDocument doc = new XWPFDocument();
		try (OutputStream os = new FileOutputStream("Word/JavatpointAlingnin.docx")) {
			XWPFParagraph paragraph = doc.createParagraph();
			XWPFParagraph paragraph2 = doc.createParagraph();
			XWPFParagraph paragraph3 = doc.createParagraph();
			XWPFParagraph paragraph4 = doc.createParagraph();
			XWPFParagraph paragraph5 = doc.createParagraph();
			/*
			 ParagraphAlignment.RIGHT aligns the Paragraph to the Right.
			 ParagraphAlignment.LEFT aligns the Paragraph to the Left.
			 ParagraphAlignment.CENTER aligns the Paragraph to the Center.
			 */
			
			paragraph.setAlignment(ParagraphAlignment.RIGHT);
			XWPFRun run = paragraph.createRun();
			run.setText("Text is aligned right");
			
			paragraph2.setAlignment(ParagraphAlignment.LEFT);
			XWPFRun run2 = paragraph2.createRun();
			run2.setText("Text is aligned left");
			
			paragraph3.setAlignment(ParagraphAlignment.CENTER);
			XWPFRun run3 = paragraph3.createRun();
			run3.setText("Text is aligned center");
			
			paragraph4.setAlignment(ParagraphAlignment.DISTRIBUTE);
			XWPFRun run4 = paragraph4.createRun();
			XWPFRun run5 = paragraph5.createRun();
			run5.setBold(true);
			run5.setText("80");
			run4.setText("Text is aligned Justificad bycause is the example with $davidx " + run5.getText(0) + " caracterect and ddd dddddd ddddd eeee wwwwww qqqqqqqq jsdsdaksd adajdajkdaksdkdjas anbjasdkjdkadn anjnajkdnbaskdj nsajdasjkdjasdajkdakjsd jnjasndjasndjkandkjasdnkajdnjak jnajsdnajksdnkjasdnakjdakjsdakjdnasjkdnaksj jnasdjkasndjkasndkjandkjasdkjasd jadnjsadnajksdkjnjksadakjsdak jasndjadnkjasdk");
			
			String text = run4.getText(0);
            if (text != null && text.contains("$davidx")) {
                text = text.replace("$davidx", "Jimena");
                run4.setText(text, 0);
            }
			
			doc.write(os);
		
			System.out.println("ok");
			System.out.println(run5.getText(0));
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}
