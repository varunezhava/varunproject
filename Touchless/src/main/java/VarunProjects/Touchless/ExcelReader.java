package VarunProjects.Touchless;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelReader {
	
	
	public XSSFWorkbook GetWorkbook(String filepath) throws IOException 
	{
		
		XSSFWorkbook workbook = new XSSFWorkbook(); 
		
		try 
		{
			InputStream xl = new FileInputStream(filepath);
			 workbook = new XSSFWorkbook(xl); 
			
		} catch (FileNotFoundException e) 
		{
			// TODO Auto-generated catch block
			System.out.println("No file found at the give path");
			e.printStackTrace();
		}
		
		
		return workbook;
		
	}
	
	
	public Sheet GetSheet(XSSFWorkbook workbook, String sheetname)
	{
		Sheet worksheet = workbook.getSheet(sheetname);

		return worksheet;
					
	}
	
	public int GetTotalRows(Sheet worksheet)
	{
		int n = worksheet.getLastRowNum();
		
		return n;
	}
	
	public String[][] DataCollector(Sheet sheet, int x, int n)
	{
	   	String data[][] = new String[x][5];
    	for(int i=1;i<=x;i++)
    	{
    	
    		Row currow = sheet.getRow(i);
    		for (int k=0;k<n;k++)
    		{
    			data[i-1][k]=currow.getCell(k).toString();
    		}
    		
    		
    	
    	}
    	return data;
    
		
	}
	

}
