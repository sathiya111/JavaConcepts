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

public class WriteDataFromexcel {

	public static void main(String[] args) throws IOException {
		
		//File f=new File("C:\\Users\\SATHIYANARAYANAN M\\Desktop\\SathiyaDetails.xlsx");
		String s="C:\\Users\\SATHIYANARAYANAN M\\Desktop\\SathiyaDetails.xlsx";
		FileInputStream fin=new FileInputStream(s);
		Workbook wb=new XSSFWorkbook(fin);
		Sheet sheet=wb.getSheet("sheet1");
		Row row = sheet.getRow(4);
		
		if(row==null) {
			Cell createCell = sheet.createRow(4).createCell(10);
			createCell.setCellValue("MCA");
		}else {
			Cell cell = row.getCell(10);
			if(cell==null) {
				row.createCell(10).setCellValue("MCA");
			}else {
				cell.setCellValue("MCA");
			}
		}
		FileOutputStream fout=new FileOutputStream(s);
		wb.write(fout);
		fout.close();
		System.out.println("Done");
	}

}
