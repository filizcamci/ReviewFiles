package workWithFiles;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FromMenuToOrder {

	public static void main(String[] args) throws IOException {
		XSSFWorkbook workBook;
		XSSFSheet sheet;
		Row row;
		List<String> fdd=new ArrayList<>();
		List<List<String>> values=new ArrayList<>();
		String path = "C:\\Users\\filiz\\Dropbox\\test\\ReviewFiles\\Menus.xlsx";
		FileInputStream fis = new FileInputStream(path);
		workBook = new XSSFWorkbook(fis);
		sheet = workBook.getSheetAt(0);
		int colomnCount = sheet.getRow(0).getLastCellNum();
		int rowCount=sheet.getPhysicalNumberOfRows();
		for(int i=0;i<rowCount;i++) {
			List<String> value=new ArrayList<>();
			row=sheet.getRow(i);
			for(int j=0; j<colomnCount;j++) {
				Cell cell=row.getCell(j);
				if(i==0)
					fdd.add(cell.toString());
				
				value.add(cell.toString());
				System.out.print(cell+"-");
			}
			values.add(value);
			System.out.println();
		}
		
		
		fis.close();
		workBook.close();
		workBook=new XSSFWorkbook();
		sheet=workBook.createSheet("menu");
		Row row0=sheet.createRow(0);
		for(int i=0;i<colomnCount;i++) {
			Cell cell0=row0.createCell(i);
			cell0.setCellValue(fdd.get(i));
		}
		for(int j=1;j<rowCount;j++) {
			row=sheet.createRow(j);
			for(int k=0;k<colomnCount;k++) {
				Cell cell=row.createCell(k);
				cell.setCellValue(values.get(j).get(k));
			}
		}
		// Resize all columns to fit the content size
		  for(int i = 0; i < colomnCount; i++) {
			  sheet.autoSizeColumn(i);
	        }
		try(FileOutputStream out=new FileOutputStream(new File("newMenu.xlsx"))){
			workBook.write(out);
		}
	}

}
