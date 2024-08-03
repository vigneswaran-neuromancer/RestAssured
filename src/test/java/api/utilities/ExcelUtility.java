package api.utilities;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtility {
	
	public FileInputStream FileInput;
	public FileOutputStream FileOutput;
	public XSSFWorkbook wrkbook;
	public XSSFSheet sheet;
	public XSSFRow row;
	public XSSFCell cell;
	public CellStyle style;
	String FilePath;
	
	public ExcelUtility(String FilePath) {
		this.FilePath=FilePath;
	}
	
	public int getRowCount(String sheetName) throws IOException
	{
		FileInput=new FileInputStream(FilePath);
		wrkbook=new XSSFWorkbook(FileInput);
		sheet=wrkbook.getSheet(sheetName);
		int rowCount=sheet.getLastRowNum();
		wrkbook.close();
		FileInput.close();
		return rowCount;
	}
	
	public int getCellCount(String sheetName, int rowNum) throws IOException
	{
		FileInput=new FileInputStream(FilePath);
		wrkbook=new XSSFWorkbook(FileInput);
		sheet=wrkbook.getSheet(sheetName);
		row=sheet.getRow(rowNum);
		int cellCount=row.getLastCellNum();
		wrkbook.close();
		FileInput.close();
		return cellCount;
	}
	
	public String getCellData(String sheetName,int rowNum,int column) throws IOException
	{
		FileInput=new FileInputStream(FilePath);
		wrkbook=new XSSFWorkbook(FileInput);
		sheet=wrkbook.getSheet(sheetName);
		row=sheet.getRow(rowNum);
		cell=row.getCell(column);
		DataFormatter formatter = new DataFormatter();
		String data;
		try {
			data=formatter.formatCellValue(cell);
		}
		catch(Exception e)
		{
			data="";
		}
		wrkbook.close();
		FileInput.close();
		return data;
	}
	

}
