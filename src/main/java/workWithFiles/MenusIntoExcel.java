package workWithFiles;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MenusIntoExcel {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		List<String> menuNames=Arrays.asList("Food","Drink","Desert");
		List<Menu> menuList=new ArrayList<>();
		Menu menu1 = new Menu("breakfast");
		Menu menu2=new Menu("lunch");
		Menu menu3=new Menu("dinner");
		Menu menu4=new Menu("snack");
		Menu menu5 = new Menu("breakfast");
		Menu menu6=new Menu("lunch");
		//System.out.println(menu1.drink);
		menuList.add(menu1);
		menuList.add(menu2);
		menuList.add(menu3);
		menuList.add(menu4);
		menuList.add(menu5);
		menuList.add(menu6);
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Today's Menu");
		Row row=sheet.createRow(0);
		for(int i=0;i<3;i++) {
			Cell cell=row.createCell(i);
			cell.setCellValue(menuNames.get(i));
		}
		for(int j=1;j<menuList.size()+1;j++) {
			Row rowj=sheet.createRow(j);
			rowj.createCell(0).setCellValue(menuList.get(j-1).food);
			rowj.createCell(1).setCellValue(menuList.get(j-1).drink);
			rowj.createCell(2).setCellValue(menuList.get(j-1).desert);
			
		}
		// Resize all columns to fit the content size
		  for(int i = 0; i < 3; i++) {
			 sheet.autoSizeColumn(i);
	        }
		try (FileOutputStream outputStream = new FileOutputStream(new File("Menus.xlsx"))) {
			workbook.write(outputStream);
		}
	}

}

class Menu {
	String food;
	String drink;
	String desert;

	public Menu(String meal) {
		if (meal.equalsIgnoreCase("Breakfast")) {
			this.food = "egg";
			this.drink = "tea";
			this.desert = "cake";
		} else if (meal.equalsIgnoreCase("Lunch")) {
			this.food = "pizza";
			this.drink = "apple juice";
			this.desert = "chocolate";
		} else if (meal.equalsIgnoreCase("Dinner")) {
			this.food = "rice and bean";
			this.drink = "yogurt drink";
			this.desert = "ice-cream";
		} else {
			this.food = "fruit";
			this.drink = "milk";
			this.desert = "non";
		}
	}
}