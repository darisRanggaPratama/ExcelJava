package ExcelDBase;

import org.junit.jupiter.api.Assertions;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;

public class App {
    static String jdbcUrl, username, password;

    public static void main(String[] args) {
        testConnect();
    }

    static void testConnect() {
        jdbcUrl = "jdbc:mysql://localhost:3306/autoclub";
        username = "rangga";
        password = "rangga";

        try {
            Connection connection = DriverManager.getConnection(jdbcUrl, username, password);
            System.out.println("Connected \n" + connection);
        } catch (SQLException exception) {
            Assertions.fail(exception);

        }
    }
}
