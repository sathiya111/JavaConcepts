package dataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ReadingDataFromExcel {
	// xlsx=XSSFWorkbook
	//xlx-HSSFWorkbook
	public static void main(String[] args) throws IOException {
		File f=new File("C:\\Users\\SATHIYANARAYANAN M\\Desktop\\EmployeeProfile.xlsx");
		FileInputStream fin=new FileInputStream(f);
		XSSFWorkbook wb = new XSSFWorkbook(fin);
		XSSFSheet sheet = wb.getSheetAt(0);//wb.getSheet("sheet1")
		
		int rows = sheet.getPhysicalNumberOfRows();
		System.out.println("Number of rows "+rows);
		for (int i = 0; i <rows ; i++) {
			Row row = sheet.getRow(i);
			int numberOfCells = row.getPhysicalNumberOfCells();
			String cellValue="";
			for (int j = 0; j < numberOfCells; j++) {
				Cell cell = row.getCell(j);
				if (cell.getCellTypeEnum().equals(CellType.STRING)) {
					cellValue = cell.getStringCellValue();
					System.out.println(cellValue);
				}
				else if (cell.getCellTypeEnum().equals(CellType.NUMERIC)) {
					double d = cell.getNumericCellValue();
					long l=(long) d;
					cellValue = String.valueOf(l);
					System.out.println(cellValue);
				}
			}
		}
		wb.close();
	}

}
