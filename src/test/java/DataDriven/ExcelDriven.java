package DataDriven;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDriven {

	
	
	public static void main(String[] args) throws IOException {

		// If you want to use excel data then create a object of XSSFWorkbook class and
		// pass the excel path as
		// FileInputStream object

		FileInputStream fis = new FileInputStream(
				"C:\\Users\\dibya\\OneDrive\\Desktop\\Desktop Items\\ExcelDriven.xlsx");

		XSSFWorkbook workbook = new XSSFWorkbook(fis);

		// Go to your desire sheet
		int sheetCount = workbook.getNumberOfSheets();
		for (int i = 0; i < sheetCount; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("TestData")) {
				XSSFSheet sheet = workbook.getSheetAt(i);

				// Step 1: Identify the testCase column by scanning the entire 1st row
				// This below command will move to iterate each of the row into the excel
				Iterator<Row> rows = sheet.iterator();
				Row firstrow = rows.next(); // This command will go to first row

				Iterator<Cell> ce = firstrow.cellIterator(); // Access to the first Cell

				int k = 0;
				int column = 0;
				while (ce.hasNext()) {
					Cell value = ce.next();
					if (value.getStringCellValue().equalsIgnoreCase("TestCases")) {
						column=k;
					}
					k++;
				}
				System.out.println(column);
				//Step 2: Once the column is identified then scan entire test column to identify purchase test case row
				
				while (rows.hasNext()) 
				{
					Row r=rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase("Purchase"))
					{
						// Step 3: after you grab purchase testcase row= pull all the data from that row and feed into test.
						Iterator<Cell> cv=r.cellIterator();
						
						while(cv.hasNext())
						{
							System.out.println(cv.next().getStringCellValue());
						}
					}
				}
			}
			
		}

	}

}
