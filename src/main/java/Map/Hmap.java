package Map;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;



public class Hmap {
	
	public static List<LinkedHashMap<String, String>> getExcelDataAsMap(String excelFileName, String sheetName) throws EncryptedDocumentException, IOException {
		// Create a Workbook
		Workbook wb = WorkbookFactory.create(new File("D:\\wrk\\newsp\\TestPOI\\ExcelFiles\\test.xlsx"));
		// Get sheet with the given name "Sheet1"
		Sheet s = wb.getSheet(sheetName);
		// Initialized an empty List which retain order
		List<LinkedHashMap<String, String>> dataList = new ArrayList<LinkedHashMap<String, String>>();
		// Get data set count which will be equal to cell counts of any row
		int ColCount = s.getRow(0).getPhysicalNumberOfCells();
		// Skipping first column as it is field names
		for (int i = 0; i < ColCount; i++) {
		
			LinkedHashMap<String, String> data  = new LinkedHashMap<String, String>();
		
			int rowCount = s.getPhysicalNumberOfRows();
			System.out.println(rowCount);
			
			for(int j = 0; j<rowCount ; j++) {
				Cell c = s.getRow(i).getCell(j);
				Cell cc = s.getRow(1).getCell(j);
				System.out.println(i);
				System.out.println(j);
				// Row r = r.getCell(j);
			String cellnum =String.valueOf(c);	// Field name is constant as 0th index
			String cellval =String.valueOf(cc);
				//String fieldName = r.getCell(0).getStringCellValue();
				//String fieldValue = r.getCell(i).getStringCellValue();
				// Add data in map
				data.put(cellnum,cellval);
				System.out.println(cellnum);
				System.out.println(cellval);
				
			}
			// Add map to list after each iteration
			dataList.add(data);
			
		}
		return dataList;
		
		  //Row row1 = s.getRow(0);       
		  /*   int colNum = -1;
		     for (int i = 0 ;i<=row1.getLastCellNum();i++){
		         Cell cell1 = row1.getCell(i);
		         String cellValue1 = cell1.getStringCellValue();
		         if ("Employee".equals(cellValue1)){
		              colNum = i ;
		              break;}
		        }
*/
		
	}
 
	public static void main(String[] args) throws EncryptedDocumentException, IOException {
 
List<LinkedHashMap<String, String>> mapDataList = getExcelDataAsMap("ExcelDataToReadInListOfMap","Sheet1");
		
		for(int k = 0; k<mapDataList.size() ; k++)
		{
			System.out.println("Data Set "+ (k+1) +" : ");
			for(String s: mapDataList.get(k).keySet())
			{
				System.out.println("Value of "+s +" is  : "+mapDataList.get(k).get(s));
			}
			System.out.println("========================================================");
		}
	}
		
	}


