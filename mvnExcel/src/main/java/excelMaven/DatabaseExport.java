package excelMaven;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

public class DatabaseExport {
    static XSSFWorkbook wbook;
    static XSSFSheet wsheet;
    static XSSFRow row;
    static XSSFCell cell;
    static Connection connection;
    static Statement statement;
    static String sql;
    static ResultSet resultSet;
    static FileOutputStream result;
    static File file;

    public static void ExcelResult() throws SQLException, IOException {
        connection = DatabaseConnect.getDataSource().getConnection();
        statement = connection.createStatement();

        sql = """
                SELECT * FROM customer;
                """;

        resultSet = statement.executeQuery(sql);

        wbook = new XSSFWorkbook();
        wsheet = wbook.createSheet("email db");

        row = wsheet.createRow(1);

        cell = row.createCell(1);
        cell.setCellValue("ID");
        cell = row.createCell(2);
        cell.setCellValue("EMAIL");
        cell = row.createCell(3);
        cell.setCellValue("NAME");

        int i = 2;
        while (resultSet.next()) {
            row = wsheet.createRow(i);
            cell = row.createCell(1);
            cell.setCellValue(resultSet.getString("id"));
            cell = row.createCell(2);
            cell.setCellValue(resultSet.getString("email"));
            cell = row.createCell(3);
            cell.setCellValue(resultSet.getString("name"));
            i++;
        }

        file = new File("./src/main/java/checking/Email.xlsx");
        result = new FileOutputStream(file);
        wbook.write(result);

        result.close();
        resultSet.close();
        statement.close();
        connection.close();

        System.out.println("File created");
    }
}
