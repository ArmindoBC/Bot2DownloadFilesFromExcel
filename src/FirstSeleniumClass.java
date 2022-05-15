
import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Vector;
import java.io.IOException;  
import org.apache.poi.hssf.usermodel.HSSFSheet;  
import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.FormulaEvaluator;  
import org.apache.poi.ss.usermodel.Row;  

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;




public class FirstSeleniumClass {
	
	WebDriver driver; 

	public void launchBrowser() {
		
		Vector<String> websites  = readExcel(); 
		
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\barbo.DESKTOP-EOMSI52\\Downloads\\chromedriver_win32\\chromedriver.exe");
		driver = new ChromeDriver();
		
		HashMap<String,Object> chromePrefs = new HashMap<String, Object>();
		chromePrefs.put("plugins.always_open_pdf_externally", true);
		chromePrefs.put("download.default_directory","C:\\Users\\barbo.DESKTOP-EOMSI52\\Downloads\\eic");
		
		ChromeOptions options = new ChromeOptions();
		options.setExperimentalOption("prefs", chromePrefs);

		driver = new ChromeDriver(options);
		

	
		for(int i = 0; i < websites.size(); ++i) {
			String aux = websites.get(i);
			driver.get(aux);
		}
		
	}
	
	public static Vector<String> readExcel() {
		
		Vector<String> websites = new Vector<String>();
		
		try
        {
            FileInputStream file = new FileInputStream(new File("excel.xlsx"));
 
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);
 
            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
 
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) 
            {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                int aux = 0; 
                while (cellIterator.hasNext()) 
                {	
                    Cell cell = cellIterator.next();
                	if(aux == 12) {
               
                    //Check the cell type and format accordingly
	                    switch (cell.getCellType()) 
	                    {
	                        case NUMERIC:
	                            System.out.print(cell.getNumericCellValue() + "t");
	                            break;
	                        case STRING:
	                        	
	                            System.out.print(cell.getStringCellValue() + "t");
	                            websites.add(cell.getStringCellValue());
	                            break;
						default:
							break;
	                    }
                	}
                    
                    
                    ++aux; 
                }
                System.out.println("");
            }
            file.close();
        } 
        catch (Exception e) 
        {
            e.printStackTrace();
        }
		
		return websites; 
		
		
	}
	
	
	public static void main(String[] args) {
		
	
		
		
		FirstSeleniumClass object = new FirstSeleniumClass(); 
		
		object.launchBrowser(); 
		
		
		
	}
	


}
