package databaseReview;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GetConnected {
	static Connection conn;
	static Statement st;
	static Statement st2;
	static ResultSet rs1;
	static ResultSet rs2;

	public static void main(String[] args) throws SQLException, IOException {
		List<List<String>> dataList1 = new ArrayList<>();
		List<List<String>> dataList2 = new ArrayList<>();
		List<String> colNames = new ArrayList<>();
		String dbURL = "jdbc:postgresql://room-reservation-qa.cxvqfpt4mc2y.us-east-1.rds.amazonaws.com:5432/room_reservation_qa";
		String user = "qa_user";
		String password = "Cybertek11!";
		conn = DriverManager.getConnection(dbURL, user, password);
		if (conn != null)
			System.out.println("connection established!");
		st = conn.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
		st2 = conn.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
		rs1 = st.executeQuery("select firstname, lastname from users;");
		rs2 = st2.executeQuery("select role from users where role='teacher';");
		ResultSetMetaData rsmd1 = rs1.getMetaData();
		ResultSetMetaData rsmd2 = rs2.getMetaData();
		// System.out.println(rsmd1.getColumnCount());
		rs1.next();
		int colCount = rsmd1.getColumnCount();
		int colCount2 = rsmd2.getColumnCount();
		for (int i = 1; i <= colCount; i++) {
			colNames.add(rsmd1.getColumnName(i));
			//System.out.print(rsmd1.getColumnName(i) + " ");
		}
		System.out.println();
		while (rs1.next()) {
			List<String> data = new ArrayList<>();
			for (int j = 1; j <= colCount; j++) {
				data.add(rs1.getObject(j).toString());
				//System.out.print(rs1.getObject(j) + " ");
			}
			dataList1.add(data);
			//System.out.println();
		}
		rs1.last();
		int rowCount = rs1.getRow();
		
		// System.out.println(rowCount);
		// System.out.println(dataList);

		while (rs2.next()) {
			List<String> data2 = new ArrayList<>();
			for (int j = 0; j < colCount2; j++) {
				data2.add(rs2.getObject(j+1).toString());
			}
			dataList2.add(data2);
			
		}

		
		rs2.last();
		int rowCount2 = rs2.getRow();
		//System.out.println(rowCount2);

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("user names");
		Row row;
		for (int i = 0; i < rowCount - 1; i++) {
			row = sheet.createRow(i);
			for (int j = 0; j < colCount; j++) {
				Cell cell = row.createCell(j);
				if (i == 0) {
					cell.setCellValue(colNames.get(j));
				} else {
					cell.setCellValue(dataList1.get(i).get(j));
				}
			}

		}
		// Resize all columns to fit the content size
		for (int i = 0; i < 3; i++) {
			sheet.autoSizeColumn(i);
		}

		XSSFSheet sheet2 = workbook.createSheet("roles");
		System.out.println("================");
		//System.out.println(dataList2);
		for (int i = 0; i < rowCount2-1; i++) {
			row = sheet2.createRow(i);
			for (int j = 0; j < colCount2; j++) {
				Cell cell = row.createCell(j);
				if (i == 0)
					cell.setCellValue("role");
				else {
					//System.out.print("i:"+i+" j:"+j+dataList2);
					cell.setCellValue(dataList2.get(i).get(j));
				}

			}
		}

		try (FileOutputStream out = new FileOutputStream(
				new File("C:\\Users\\filiz\\Dropbox\\test\\ReviewFiles\\users.xlsx"))) {
			workbook.write(out);
		}
		workbook.close();
		st.close();
		st2.close();
		conn.close();
	}

}
