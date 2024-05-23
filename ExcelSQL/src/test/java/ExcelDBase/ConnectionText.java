package ExcelDBase;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.sql.Connection;
import java.sql.Driver;
import java.sql.DriverManager;
import java.sql.SQLException;

public class ConnectionText {
    @BeforeAll
    static void beforeAll() {
        try {
            Driver mysqlDriver = new com.mysql.cj.jdbc.Driver();
            DriverManager.registerDriver(mysqlDriver);
        } catch (SQLException e) {
            e.printStackTrace();
        }
    }

    @Test
    void testConnection() {
        String jdbcUrl, username, password;
        jdbcUrl = "jdbc:mysql://localhost:3306/java_database?serverTimezone=Asia/Jakarta";
        username = "rangga";
        password = "rangga";

        try {
            Connection connection = DriverManager.getConnection(jdbcUrl, username, password);
            System.out.println("Connected \n" + connection);
            connection.close();
            System.out.println("DisConnected \n" + connection);
        } catch (SQLException e) {
            System.out.println("Disconnected: \n");
            Assertions.fail(e);
        }
    }
}
