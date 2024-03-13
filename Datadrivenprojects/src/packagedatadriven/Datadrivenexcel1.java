package packagedatadriven;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Datadrivenexcel1 {
	XSSFWorkbook wb=new XSSFWorkbook();//global variable declaration
	XSSFSheet sheet=wb.createSheet();//global variable declaration	
	// method1 that for excle path
public Datadrivenexcel1(String excelPath) throws IOException {
	File filepath = new File("C:\\Excel Book\\sheet3.xlsx");             
	FileInputStream fis = new FileInputStream(filepath);// file path store here 
	wb = new XSSFWorkbook(fis);//location give to wb
	
}

public String getData(int sheetnumber, int row, int column)
{
	sheet = wb.getSheetAt(sheetnumber);
	String data = sheet.getRow(row).getCell(column).getStringCellValue();
	return data;
	
}

public int getRowCount(int sheetIndex) 
{
	int row = wb.getSheetAt(sheetIndex).getLastRowNum();
	row = row + 1;
	return row;

}
}
