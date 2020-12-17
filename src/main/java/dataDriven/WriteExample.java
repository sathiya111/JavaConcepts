package dataDriven;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExample {
public static void main(String[] args) throws IOException {
	String s="C:\\Users\\SATHIYANARAYANAN M\\Desktop\\SathiyaDetails.xlsx";
	FileInputStream fin=new FileInputStream(s);
	Workbook wb=new XSSFWorkbook(fin);
	Sheet sheet=wb.getSheet("sheet1");
	Row row = sheet.getRow(10);
	if (row==null) {
		Cell createCell = sheet.createRow(10).createCell(10);
		createCell.setCellValue("Ramsathiya");
		
	}else {
		Cell cell = row.getCell(10);
		if (cell==null) {
			row.createCell(10).setCellValue("Ramsathiya");
			cell.setCellValue("ramsathiya");
		}
	}
	FileOutputStream fos=new FileOutputStream(s);
	wb.write(fos);
	fos.close();
	System.out.println("Done");
}
}
