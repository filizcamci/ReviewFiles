package apiReview;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import io.restassured.RestAssured;
import io.restassured.path.json.JsonPath;
import io.restassured.response.Response;

public class ApiToExcel {

	public static void main(String[] args) throws IOException {
		
		Response res = RestAssured.given().when().get("https://api.got.show/api/characters/");
		JsonPath jp = res.jsonPath();
		//System.out.println(jp.prettyPrint());
		List<String> houseList=jp.getList("house");
		//System.out.print(houseList);
		
		for(String house:houseList) {
			System.out.println(house);
		}
		Set<String> houseSet=new HashSet<>(houseList);
		System.out.println(houseList.size());
		System.out.println(houseSet.size());
		//https://api.got.show/api/characters/
		XSSFWorkbook workbook;
		XSSFSheet sheet;
		Row row;
		int rowCount=houseList.size();
		workbook=new XSSFWorkbook();
		sheet=workbook.createSheet("House");
		for(int i=0;i<rowCount;i++) {
		row=sheet.createRow(i);
		Cell cell=row.createCell(0);
		if(houseList.get(i)!=null)
			cell.setCellValue(houseList.get(i));
		}
		
		// Resize all columns to fit the content size
				for (int i = 0; i < 3; i++) {
					sheet.autoSizeColumn(i);
				}
				
				try (FileOutputStream out = new FileOutputStream(
						new File("C:\\Users\\filiz\\Dropbox\\test\\ReviewFiles\\GOT.xlsx"))) {
					workbook.write(out);
	}
				workbook.close();
	}

}
