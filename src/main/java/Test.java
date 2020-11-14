import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Test {
	 public void readExcel(String filePath,String fileName,String sheetName) throws IOException{

		   

		    File file =    new File(filePath+"\\"+fileName);
		    FileInputStream inputStream = new FileInputStream(file);
		    Workbook wb = null;
		    //Find the file extension by splitting file name in substring  and getting only extension name
		    String fileExtensionName = fileName.substring(fileName.indexOf("."));
		    //Check condition if the file is xlsx file
		    if(fileExtensionName.equals(".xlsx")){
		    //If it is xlsx file then create object of XSSFWorkbook class
		    	wb = new XSSFWorkbook(inputStream);
		    }
		    else if(fileExtensionName.equals(".xls")){
		        //If it is xls file then create object of HSSFWorkbook class
		    	wb = new HSSFWorkbook(inputStream);
		    }
		    //Read sheet inside the workbook by its name
		    Sheet sheet = wb.getSheet(sheetName);
		    //Find number of rows in excel file
		    int rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum();
		    //Create a loop over all the rows of excel file to read it
		    for (int i = 0; i < rowCount+1; i++) {
		        Row row = sheet.getRow(i);
		        //Create a loop to print cell values in a row
		        for (int j = 0; j < row.getLastCellNum(); j++) {
		            //Print Excel data in console
		            System.out.print(row.equals("A1"));
		        }
		        System.out.println();
		    } 
		    }  
		    //Main function is calling readExcel function to read data from excel file
		    public static void main(String args[]) throws IOException{
				  Test objExcelFile = new Test();			  
		    String filePath = "D:\\wrk\\newsp\\TestPOI\\ExcelFiles";
		    objExcelFile.readExcel(filePath,"test.xlsx","Sheet1");
		    }

		}