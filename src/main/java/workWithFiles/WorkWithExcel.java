package workWithFiles;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

//more information: https://www.callicoder.com/java-write-excel-file-apache-poi/
public class WorkWithExcel {
	static XSSFWorkbook workbook;
	static XSSFSheet worksheet;
	static Row row;

	public static void main(String[] args) throws Exception {

		//System.out.println(getDataFromExcel("C:\\Users\\filiz\\Downloads\\MOCK_DATA (3).xlsx"));
		writeIntoExcel("C:\\Users\\filiz\\Dropbox\\test\\ReviewFiles\\Customers.xlsx");

	}

	public static List<Map<String, String>> getDataFromExcel(String path) throws Exception {
		List<Map<String, String>> dataList = new ArrayList<>();

		List<String> colomnNames = new ArrayList<>();
		FileInputStream file = new FileInputStream(path);
		workbook = new XSSFWorkbook(file);
		// worksheet=workbook.getSheet("data");
		worksheet = workbook.getSheetAt(0);

		int colomnCount = worksheet.getRow(0).getLastCellNum();
		int rowCount = worksheet.getPhysicalNumberOfRows();
		for (int k = 0; k < colomnCount; k++) {
			colomnNames.add(worksheet.getRow(0).getCell(k).toString());
		}
		for (int i = 1; i < rowCount; i++) {
			row = worksheet.getRow(i);
			Map<String, String> data = new TreeMap<>();
			for (int j = 0; j < colomnCount; j++) {
				Cell cell = row.getCell(j);
				data.put(colomnNames.get(j), cell.toString());
			}
			dataList.add(data);

		}
		for (int n = 0; n < dataList.size(); n++) {
			System.out.println(dataList.get(n));
		}
		return dataList;
	}

	public static void writeIntoExcel(String path) throws Exception {
		List<Map<String, String>> dataList = getDataFromExcel("C:\\Users\\filiz\\Downloads\\MOCK_DATA (3).xlsx");
		Set<String> keyset = dataList.get(0).keySet();
		List<String> keyList = new ArrayList<>(keyset);
		//System.out.println(keyList);
		int columnCount = keyList.size();
		int rowCount = dataList.size();
		workbook = new XSSFWorkbook();
		worksheet = workbook.createSheet("customer data");
		Row row0 = worksheet.createRow(0);
		for (int i = 0; i < columnCount; i++) {
			Cell cell = row0.createCell(i);
			cell.setCellValue(keyList.get(i));
		}
		for (int j = 1; j < rowCount+1; j++) {
			Row rowj = worksheet.createRow(j);
			for (int i = 0; i < columnCount; i++) {
				Cell cell = rowj.createCell(i);
				cell.setCellValue(dataList.get(j-1).get(keyList.get(i)));
			}
		}
		
		// Resize all columns to fit the content size
		  for(int i = 0; i < columnCount; i++) {
			  worksheet.autoSizeColumn(i);
	        }
		try (FileOutputStream outputStream = new FileOutputStream(new File(path))) {
			workbook.write(outputStream);
		}
	}

}
