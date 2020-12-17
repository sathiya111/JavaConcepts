package dataDriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import javax.imageio.stream.FileImageInputStream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookType;

public class DataDrivenExample {
public static void main(String[] args) throws FileNotFoundException, IOException {
	File f=new File("C:\\Users\\SATHIYANARAYANAN M\\Desktop\\SathiyaDetails.xlsx");
	FileInputStream fin=new FileInputStream(f);
	  XSSFWorkbook wb=new XSSFWorkbook(fin);
	 XSSFSheet sheet = wb.getSheetAt(0);
	 int rows = sheet.getPhysicalNumberOfRows();
	 for (int i = 0; i < rows; i++) {
		XSSFRow row = sheet.getRow(i);
		int cells = row.getPhysicalNumberOfCells();
		for (int j = 0; j < cells; j++) {
			XSSFCell cell = row.getCell(j);
			String cellValue="";
			if (cell.getCellTypeEnum().equals(CellType.STRING)) {
				cellValue = cell.getStringCellValue();
				System.out.println(cellValue);
			}else if (cell.getCellTypeEnum().equals(CellType.NUMERIC)) {
			double d=	cell.getNumericCellValue();
			long l=(long) d;
			String.valueOf(d);
			System.out.println(cellValue);
			}
		}
		wb.close();
	}
}
}
