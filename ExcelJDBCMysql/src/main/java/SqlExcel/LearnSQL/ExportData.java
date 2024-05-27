package SqlExcel.LearnSQL;

import SqlExcel.DatabaseConnect;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import java.sql.*;

public class ExportData {
    static XSSFWorkbook wbook;
    static XSSFSheet wsheet;
    static XSSFRow row;
    static XSSFCell cell;
    static XSSFFont font;
    static XSSFCellStyle style;
    static Connection connection;
    static PreparedStatement preState;
    static String sql;
    static ResultSet resultSet;
    static FileOutputStream output;
    static File file;

    public static void CreateExcel() throws SQLException, IOException {
        // Get connection
        connection = DatabaseConnect.getDataSource().getConnection();

        // Query Statement
        sql = """
                SELECT * FROM customer;
                """;

        // Run query
        preState = connection.prepareStatement(sql);
        resultSet = preState.executeQuery();

        // Create worksheet
        wbook = new XSSFWorkbook();
        wsheet = wbook.createSheet("email db");

        // Create a new font and alter it
        font = wbook.createFont();
        font.setFontHeightInPoints((short) 14);
        font.setFontName("COMIC SANS MS");
        font.setItalic(true);
        font.setBold(true);
        font.setColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());

        // Set font style
        style = wbook.createCellStyle();
        style.setFont(font);

        // Create row to be written
        row = wsheet.createRow(0);
        cell = row.createCell(0);
        cell.setCellValue("NO");
        cell.setCellStyle(style);
        cell = row.createCell(1);
        cell.setCellValue("ID");
        cell.setCellStyle(style);
        cell = row.createCell(2);
        cell.setCellValue("EMAIL");
        cell.setCellStyle(style);
        cell = row.createCell(3);
        cell.setCellValue("NAME");
        cell.setCellStyle(style);


        int i = 1;
        while (resultSet.next()) {
            row = wsheet.createRow(i);
            cell = row.createCell(0);
            cell.setCellValue(i);
            cell = row.createCell(1);
            cell.setCellValue(resultSet.getString("id"));
            cell = row.createCell(2);
            cell.setCellValue(resultSet.getString("email"));
            cell = row.createCell(3);
            cell.setCellValue(resultSet.getString("name"));
            i++;
        }

        // Create Excel file to be written
        file = new File("./src/main/java/created/Email.xlsx");
        output = new FileOutputStream(file);
        wbook.write(output);

        // Closing File, ResulSet, SQL statement & connection
        output.close();
        resultSet.close();
        preState.close();
        connection.close();

        System.out.println("\nFile created");
    }
}
