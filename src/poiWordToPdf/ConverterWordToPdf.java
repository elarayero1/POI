package poiWordToPdf;
/* @autor: jdvelasquez
@ feha: 15 de abr. de 2022
https://developrogramming.com/convertir-de-word-a-pdf-con-java/
https://localcoder.org/trying-to-make-simple-pdf-document-with-apache-poi
https://programmerclick.com/article/91041208944/
*/

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyles;

import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;

public class ConverterWordToPdf {
	public static void main(String[] args) {
		XWPFDocument doc = new XWPFDocument();
		try (OutputStream os = new FileOutputStream("Word/JavatpointParagraph2.doc")) {
			XWPFParagraph paragraph = doc.createParagraph();
			XWPFRun run = paragraph.createRun();
	          run.setBold(true);  
	          run.setItalic(true);  
	          run.setText("This text is Bold and have Italic style"); 
			run.setText("Hello, This is javatpoint. This paragraph is written " + "by using XWPFParagrah.");
			doc.write(os);
			System.out.println("creado");
			
			File archivoWord = new File("Word/JavatpointParagraph2.doc");
	        File archivoPDF = new File("Word/JavatpointParagraph2.pdf");
	        
	        XWPFDocument document = leerDocx(archivoWord);
	        
	        document.createStyles();
	        
	        // Se convierte el contenido del fichero Word a PDF
	        if (convertirPDF(archivoPDF, document)) {
	            // Mostramos mensaje de éxito
	            System.out.println("El fichero de Word se ha convertido a PDF con éxito.");
	        } else {
	            System.out.println("ERROR: El fichero de Word NO se ha convertido a PDF.");
	        }
			
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}
	
	public static XWPFDocument leerDocx(File archivoWord) {
		 
        XWPFDocument documentoWord = null;
 
        try {
            // Se prepara el archivo para su tratamiento
            InputStream texto = new FileInputStream(archivoWord);
           
            // Creamos documento especial POI para su posterior conversión
            documentoWord = new XWPFDocument(texto);
           // texto.
            
        } catch (IOException e) {
            System.out.println("Error leyendo el fichero de Word");
            e.printStackTrace();
        }
        return documentoWord;
 
    }
 
    public static boolean convertirPDF(File archivoPDF, XWPFDocument documentWord) {
 
        boolean exito;
 
        try {
            OutputStream out = new FileOutputStream(archivoPDF);
            
            XWPFDocument document = new XWPFDocument();
            
            XWPFParagraph run2 = document.createParagraph();
            
            // there must be a styles document, even if it is empty
            XWPFStyles styles = document.createStyles();
            
              run2.setAlignment(ParagraphAlignment.CENTER);
            
           // document.setParagraph(paragraph, pos);
            
            /*XWPFParagraph paragraph = doc.createParagraph();
			XWPFRun run = paragraph.createRun();
	          run.setBold(true);  
	          run.setItalic(true);  
	          run.setText("This text is Bold and have Italic style");  */
            
          PdfOptions options = PdfOptions.create();
          //options.
            PdfConverter.getInstance().convert(documentWord, out, options);
            
 
            exito = true;
        } catch (IOException e) {
            exito = false;
            System.out.println("Error creando el fichero PDF");
            e.printStackTrace();
        }
 
        return exito;
 
    }
}
