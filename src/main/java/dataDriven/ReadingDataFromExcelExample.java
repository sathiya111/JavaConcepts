package dataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Set;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.schemas.office.visio.x2012.main.CellType;

public class ReadingDataFromExcelExample {
	public static void main(String[] args) throws IOException {
		
	
File f=new File("C:\\Users\\SATHIYANARAYANAN M\\Desktop\\SathiyaDetails.xlsx");
FileInputStream fin=new FileInputStream(f);
  XSSFWorkbook wb=new XSSFWorkbook(fin);
 XSSFSheet sheet = wb.getSheetAt(0);
 int NumberOfRows = sheet.getPhysicalNumberOfRows();
 for (int i = 0; i < NumberOfRows; i++) {
	 Row rowcount = sheet.getRow(i);
	 int NumberOfCells = rowcount.getPhysicalNumberOfCells();
	 for (int j = 0; j < NumberOfCells; j++) {
		Cell cell = rowcount.getCell(j);
		String cellValue="";
		if (cell.getCellTypeEnum().equals(org.apache.poi.ss.usermodel.CellType.STRING)) {
			cellValue=cell.getStringCellValue();
			System.out.println(cellValue);
			
		}else if (cell.getCellTypeEnum().equals(org.apache.poi.ss.usermodel.CellType.NUMERIC)) {
			double a = cell.getNumericCellValue();
			long l=(long) a;
			cellValue=String.valueOf(l);
			System.out.println(cellValue);
		}
		wb.close();
		}
	 
	
}
}
}