package snmaddula.quicktricks.app;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * A utility class to transform given resultset into an excel workbook.
 * 
 * @author snmaddula
 *
 */
public class ResultSetToExcel {

	public static void writeToExcel(ResultSet rs, String filePath) throws SQLException, IOException {

		// Create a workbook and add the header row.
		try(Workbook book = new XSSFWorkbook()) {
			Sheet sheet = book.createSheet();
			Row header = sheet.createRow(0);
			
			// Extract columns from the ResultSet.
			ResultSetMetaData rsmd = rs.getMetaData();
			List<String> columns = new ArrayList<String>() {{
				for (int i = 1; i <= rsmd.getColumnCount(); i++) 
					add(rsmd.getColumnLabel(i));
			}};
		
			// Populate the header row with the extracted column names.
			for (int i = 0; i < columns.size(); i++) {
				Cell cell = header.createCell(i);
				cell.setCellValue(columns.get(i));
			}
			
			// Extract the rows from the ResultSet and populate the workbook.
			int rowIndex = 0;
			while (rs.next()) {
				Row row = sheet.createRow(++rowIndex);
				for (int i = 0; i < columns.size(); i++) {
					row.createCell(i).setCellValue(Objects.toString(rs.getObject(columns.get(i)), ""));
				}
			}
		
			try(FileOutputStream fos = new FileOutputStream(filePath)) {
				book.write(fos);
			}
		}
	}

}
