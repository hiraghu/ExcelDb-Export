package excel.db;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.text.Document;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

public class ExcelData {

	final static String url = "jdbc:mysql://localhost:3306/advjdb";
	final static String user_host = "root";
	final static String pwd = "1234";
	private static final String query = " select * from emp";

	public static void main(String[] args) throws Exception {

		Connection con = DriverManager.getConnection(url, user_host, pwd);
		Statement stmt = con.createStatement();
		ResultSet rs = stmt.executeQuery(query);
		

		// Create Excel workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Book1");

		// Write the column headers to the first row of the worksheet
		Row headerRow = sheet.createRow(0);
		int columnCount = rs.getMetaData().getColumnCount();
		for (int i = 1; i <= columnCount; i++) {
			Cell cell = headerRow.createCell(i - 1);
			cell.setCellValue(rs.getMetaData().getColumnName(i));
		}

		// Write data to sheet
		int rowNum = 1;
		while (rs.next()) {
			Row row = sheet.createRow(rowNum++);
			row.createCell(0).setCellValue(rs.getInt(1));
			row.createCell(1).setCellValue(rs.getString(2));

			row.createCell(2).setCellValue(rs.getString(3));
			System.out.println("inserting");
		

		}

		// Save workbook to file
		File file = new File("E:\\New folder\\Book1.xlsx");
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		out.close();
		System.out.println("EmployeeData.xlsx written successfully!");

	}
}
