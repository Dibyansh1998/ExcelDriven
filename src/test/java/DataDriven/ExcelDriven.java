package DataDriven;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDriven {

	public ArrayList<String> getData(String testcaseName) {
		ArrayList<String> testData = new ArrayList<>();

		String path = System.getProperty("user.dir") + "\\ExcelFolder\\ExcelDriven.xlsx";
		try (FileInputStream fis = new FileInputStream(path); XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

			Sheet sheet = workbook.getSheet("TestData");
			Iterator<Row> rows = sheet.iterator();

			int columnIndex = -1;

			// Find the column index for "TestCases"
			Row firstRow = rows.next();
			Iterator<Cell> cellIterator = firstRow.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				if (cell.getStringCellValue().equalsIgnoreCase("TestCases")) {
					columnIndex = cell.getColumnIndex();
					break;
				}
			}

			if (columnIndex == -1) {
				throw new RuntimeException("TestCases column not found in the Excel sheet.");
			}

			// Iterate over rows to find the matching testcaseName
			while (rows.hasNext()) {
				Row row = rows.next();
				if (row.getCell(columnIndex).getStringCellValue().equalsIgnoreCase(testcaseName)) {
					// Fetch data from the matching row
					Iterator<Cell> cellIteratorForRow = row.cellIterator();
					while (cellIteratorForRow.hasNext()) {
						Cell cell = cellIteratorForRow.next();
						if (cell.getCellType() == CellType.STRING) {
							testData.add(cell.getStringCellValue());
						} else if (cell.getCellType() == CellType.NUMERIC) {
							testData.add(String.valueOf(cell.getNumericCellValue()));
						}
					}
					break; // Exit loop after finding the matching row
				}
			}
		} catch (IOException e) {
			e.printStackTrace();
		}

		return testData;
	}
}
