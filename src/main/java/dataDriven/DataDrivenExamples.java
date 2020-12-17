package dataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.schemas.office.visio.x2012.main.CellType;

public class DataDrivenExamples {
public static void main(String[] args) throws IOException {
	File f=new File("C:\\Users\\SATHIYANARAYANAN M\\Desktop\\EmployeeProfile.xlsx");
	FileInputStream fin=new FileInputStream(f);
	Workbook wb=new XSSFWorkbook(fin);
	Sheet sheet = wb.getSheet("sheet1");
	int NumberOfRows = sheet.getPhysicalNumberOfRows();
	for (int i = 0; i < NumberOfRows; i++) {
		Row row = sheet.getRow(i);
		int NumberOfCells = row.getPhysicalNumberOfCells();
		for (int j = 0; j < NumberOfCells; j++) {
			Cell cell = row.getCell(j);
			String cellvalue="";
			if (cell.getCellTypeEnum().equals(org.apache.poi.ss.usermodel.CellType.STRING)) {
				cellvalue=cell.getStringCellValue();
				System.out.println(cellvalue);
			}else if (cell.getCellTypeEnum().equals(org.apache.poi.ss.usermodel.CellType.NUMERIC)) {
				double d = cell.getNumericCellValue();
				Long l=(long) d;
				cellvalue=String.valueOf(d);
				System.out.println(cellvalue);
			}
		}
	}
}
}
