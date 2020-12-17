package dataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.xb.xsdschema.impl.FieldDocumentImpl;

public class WriteDataExample {
public static void main(String[] args) throws IOException {
	File f= new File("C:\\\\Users\\\\SATHIYANARAYANAN M\\\\Desktop\\\\SathiyaDetails.xlsx");
	FileInputStream fin=new FileInputStream(f);
	Workbook wb=new XSSFWorkbook(fin);
	Sheet sheet = wb.getSheet("sheet1");
	Row row = sheet.getRow(13);
	if (row==null) {
		Cell createCell = sheet.createRow(13).createCell(13);
		createCell.setCellValue("FormTractor");
		
	}else {
		Cell cell = row.getCell(13);
		
		if (cell==null) {
			
			row.createCell(13).setCellValue("FormTractor");
			cell.setCellValue("FormTractor");
		}
	}
	FileOutputStream fos=new FileOutputStream(f);
	wb.write(fos);
	fos.close();
	System.out.println("Done");
	
}
}
