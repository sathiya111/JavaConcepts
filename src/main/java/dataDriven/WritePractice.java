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

public class WritePractice {
public static void main(String[] args) throws IOException {
	File f=new File("C:\\Users\\SATHIYANARAYANAN M\\Desktop\\EmployeeProfile.xlsx");
	FileInputStream fin= new FileInputStream(f);
	Workbook wb=new XSSFWorkbook(fin);
	Sheet sheet = wb.getSheet("sheet1");
	Row row = sheet.getRow(15);
	if (row==null) {
		Cell createCell = sheet.createRow(15).createCell(15);
		createCell.setCellValue("Homework");
	}else {
		Cell cell = row.getCell(15);
		if (cell==null) {
			row.createCell(15).setCellValue("Homework");
		cell.setCellValue("Homework");
		}
		
	}
	FileOutputStream fos=new FileOutputStream(f);
	wb.write(fos);
	fos.close();
	System.out.println("Successfully ");
}
}
