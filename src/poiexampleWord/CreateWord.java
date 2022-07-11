package poiexampleWord;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class CreateWord {
	public static void main(String[] args) throws FileNotFoundException, IOException {
		XWPFDocument document = new XWPFDocument();
		try (OutputStream fileOut = new FileOutputStream("Word/Javatpoint.docx")) {
			document.write(fileOut);
			System.out.println("File created");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
}