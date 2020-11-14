package next;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import javax.xml.crypto.Data;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.xmlbeans.impl.xb.xsdschema.Public;


public class ReadWriteExcelProperties {
//
	public static void main(String[] args) throws Exception {

	    File file =    new File("D:\\wrk\\newsp\\TestPOI\\ExcelFiles\\test.xlsx");
	    FileInputStream inputStream = new FileInputStream(file);


	     Workbook wb = new XSSFWorkbook(inputStream);
	     Sheet sh = wb.getSheet("Sheet1");

	    //  Find the Column number Which has column name and row number 0
	     List<LinkedHashMap<String, String>> dataList = new ArrayList<LinkedHashMap<String, String>>();
	     Row row1 = sh.getRow(0);       
	     int colNum = sh.getPhysicalNumberOfRows();
	     for (int i = 0 ;i<=colNum;i++){
	    		LinkedHashMap<String, String> data  = new LinkedHashMap<String, String>();
	         Cell cell1 = row1.getCell(i);
	         String a = String.valueOf(i);
	         String cellValue1 = cell1.getStringCellValue();
	         XSSFRow row = (XSSFRow) sh.getRow(1);
	         XSSFCell cell = row.getCell(i);
 	        if(cell != null) {
 	        //	System.out.println(cell);
 	        }
 	      
	         data.put( cell.toString(), cellValue1);
	         dataList.add(data);
	       
	     }
	     System.out.println("1 : "+dataList.get(1));
	     }}
	         //for(int rowNumber = 1; rowNumber <= sh.getLastRowNum(); rowNumber++) {
	        /*	    XSSFRow row = (XSSFRow) sh.getRow(1);

	        	    for(int columnNumber = 0; columnNumber <= row.getLastCellNum(); columnNumber++) {
	        	    	XSSFCell cell = row.getCell(columnNumber);
	        	        if(cell != null) {
	        	        //	System.out.println(cell);
	        	        }
	        	}
		      //   System.out.println(cellValue2);
	                 
	     }*/
	    

	    //  Find the Row number Which is "Y" and column number 1

	    

	   // Row r = sh.getRow(rowNum);
	   //String val = r.getCell(colNum).getStringCellValue();

	    // System.out.println(dataList.get(0));
//	     System.out.println(cellValueMaybeNull);
	     /*    System.out.println("Column number is "+colNum);
	     System.out.println("Last Column number :"+sh.getRow(0).getLastCellNum());
	     System.out.println("Value is "+val); */

	    // }//}
	
	//}
