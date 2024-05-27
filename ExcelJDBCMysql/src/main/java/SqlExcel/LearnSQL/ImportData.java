package SqlExcel.LearnSQL;

import SqlExcel.DatabaseConnect;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.jetbrains.annotations.NotNull;

import java.io.FileInputStream;
import java.io.IOException;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;

public class ImportData {
    static Connection connect;
    static FileInputStream input;
    static XSSFWorkbook wbook;
    static XSSFSheet wsheet;
    static Cell cell;
    static String filePath, id, email, name, sql;

    public static void readExcel() {
        filePath = "./src/main/java/created/Import.xlsx";

        try {
            input = new FileInputStream(filePath);
            wbook = new XSSFWorkbook(input);

            wsheet = wbook.getSheetAt(0);
            connect = DatabaseConnect.getDataSource().getConnection();

            for (Row row : wsheet) {
                // Skip title/ header (first row)
                if (row.getRowNum() == 0) {
                    continue;
                }

                if (row != null) {
                    id = getCellValue(row, 1);
                    email = getCellValue(row, 2);
                    name = getCellValue(row, 3);
                    insertData(connect, id, email, name);
                }
            }

            input.close();
            connect.close();
            System.out.println("\nData insert successfully\n");

        } catch (IOException | SQLException e) {
            System.out.println("\nFail to insert data: \n" + (e));
        }
    }

    private static String getCellValue(@NotNull Row row, int columnIndex) {
        cell = row.getCell(columnIndex);
        if (cell == null) {
            return "";
        }
        return cell.toString();
    }

    private static void insertData(@NotNull Connection connect, String id, String email, String name) throws SQLException {
        sql = """
                INSERT INTO customer(id, name, email) VALUES(?, ?, ?);
                """;
        try (PreparedStatement preState = connect.prepareStatement(sql)) {
            preState.setString(1, id);
            preState.setString(2, name);
            preState.setString(3, email);
            preState.executeUpdate();
        }
    }
}
