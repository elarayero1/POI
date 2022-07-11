package poiexampleEXCEL;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;

/* @autor: jdvelasquez
@ feha: 6 de abr. de 2022
*/
public class DataResumen {
	
	public static void main(String[] args) throws FileNotFoundException, IOException {  
		
	try (OutputStream os = new FileOutputStream("resumen.xls")) {
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet("Sheet");
		HashMap<String, Object> properties = new HashMap<String, Object>();
		// Set border around the cell
		properties.put(CellUtil.BORDER_TOP, BorderStyle.MEDIUM);
		properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.MEDIUM);
		properties.put(CellUtil.BORDER_LEFT, BorderStyle.MEDIUM);
		properties.put(CellUtil.BORDER_RIGHT, BorderStyle.MEDIUM);
		// Set color Red
		properties.put(CellUtil.TOP_BORDER_COLOR, IndexedColors.RED.getIndex());
		properties.put(CellUtil.BOTTOM_BORDER_COLOR, IndexedColors.RED.getIndex());
		properties.put(CellUtil.LEFT_BORDER_COLOR, IndexedColors.RED.getIndex());
		properties.put(CellUtil.RIGHT_BORDER_COLOR, IndexedColors.RED.getIndex());
		// Apply the borders to the cell
		Row row = sheet.createRow(1);
		Cell cell = row.createCell(1);
		CellUtil.setCellStyleProperties(cell, properties);
		// Apply the borders to a 3x3 region starting at D4
		for (int i = 1; i <= 3; i++) {
			row = sheet.createRow(i);
		
			for (int j = 1; j <= 2; j++) {
				cell = row.createCell(j);
				if(i==1 && j == 1){
					cell.setCellValue("Aniversario");
				}if(i==1 && j == 2){
					cell.setCellValue("22");
				}if(i==2 && j == 1){
					cell.setCellValue("Vacaciones");
				}if(i==2 && j == 2){
					cell.setCellValue("10");
				}if(i==3 && j == 1){
					cell.setCellValue("Reposo");
				}if(i==3 && j == 2){
					cell.setCellValue("10");
				}
				CellUtil.setCellStyleProperties(cell, properties);
			}
		}
		workbook.write(os);
		System.out.println("creado");
        }catch(Exception e) {  
            System.out.println(e.getMessage());  
        }  
    }

}
