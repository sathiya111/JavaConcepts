package dataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.schemas.office.visio.x2012.main.CellType;

public class GetdataExample {
public static void main(String[] args) throws IOException {
	File f=new File("C:\\Users\\SATHIYANARAYANAN M\\Desktop\\EmployeeProfile.xlsx");
	FileInputStream fin=new FileInputStream(f);
	Workbook wb=new XSSFWorkbook(fin);
	Sheet sheet = wb.getSheet("sheet1");
	int Rows = sheet.getPhysicalNumberOfRows();
	System.out.println("Number of Rows"+Rows);

	for (int i = 0; i < Rows; i++) {
		Row row = sheet.getRow(i);
		int numberOfCells = row.getPhysicalNumberOfCells();
		String cellValue="";

		for (int j = 0; j < numberOfCells; j++) {
			Cell cell = row.getCell(j);
			if (cell.getCellTypeEnum().equals(org.apache.poi.ss.usermodel.CellType.STRING)) {
				cellValue=cell.getStringCellValue();
				System.out.println(cellValue);
			}else if (cell.getCellTypeEnum().equals(org.apache.poi.ss.usermodel.CellType.NUMERIC)) {
				double d = cell.getNumericCellValue();
				long l=(long) d;
			cellValue=String.valueOf(l);
			System.out.println(cellValue);
			}
			
		}
		 
	
		
		
	
		
	}
}
}
