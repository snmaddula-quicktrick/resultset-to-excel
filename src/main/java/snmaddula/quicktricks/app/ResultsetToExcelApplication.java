package snmaddula.quicktricks.app;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ResultsetToExcelApplication {

	public static void main(String[] args) throws Exception {
		Connection con = DriverManager.getConnection("jdbc:h2:tcp://localhost/~/test1", "sa", "");
		if (con != null) {

			PreparedStatement ps = con.prepareStatement("SELECT * FROM ABCD");
			Workbook book = new XSSFWorkbook();

			ResultSet rs = ps.executeQuery();
			writeToExcel(book, rs, "SHEET_1");
			rs = ps.executeQuery();
			writeToExcel(book, rs, "SHEET_2");

			FileOutputStream fileOut = new FileOutputStream("poi-generated-file.xlsx");
			book.write(fileOut);
			fileOut.close();
			book.close();

		}
	}

	private static void writeToExcel(Workbook book, ResultSet rs, String sheetName) throws SQLException, IOException {

		Sheet sheet = book.createSheet(sheetName);
		Font font = book.createFont();
		font.setBold(true);
		font.setFontHeightInPoints((short) 12);

		CellStyle bold = book.createCellStyle();
		bold.setFont(font);

		Row header = sheet.createRow(0);
		List<String> columns = new ArrayList<>();

		ResultSetMetaData rsmd = rs.getMetaData();
		for (int i = 1; i <= rsmd.getColumnCount(); i++) {
			Cell cell = header.createCell(i - 1);
			cell.setCellValue(rsmd.getColumnLabel(i));
			columns.add(rsmd.getColumnLabel(i));
			cell.setCellStyle(bold);
		}
		int rowIndex = 1;
		while (rs.next()) {
			Row row = sheet.createRow(rowIndex);
			for (int i = 0; i < columns.size(); i++) {
				Cell cell = row.createCell(i);
				cell.setCellValue(Objects.toString(rs.getObject(columns.get(i)), ""));
			}
			rowIndex++;
		}

	}

}
