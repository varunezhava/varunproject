package VarunProjects.Touchless;

import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class Master 
{
    public static void main( String[] args ) throws IOException
    {
    	
    	String filepath = "E:\\Work\\Touchless\\Testcases.xlsx";
    	ExcelReader reader = new ExcelReader();
    	XSSFWorkbook book = new XSSFWorkbook();
    	
    	book = reader.GetWorkbook(filepath);
    	Sheet sheet = reader.GetSheet(book, "Sheet1");
    	int x = reader.GetTotalRows(sheet);
    
    	
    	
    	
    	
        
    }
}
