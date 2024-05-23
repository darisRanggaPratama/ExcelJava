package ExcelDBase;

import org.junit.jupiter.api.Test;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

public class InjectionSQLTest {
    @Test
    void testSQLInjection() throws SQLException {
        Connection connection = ConnectionUtil.getDataSource().getConnection();
        Statement statement = connection.createStatement();

        String username, password;
        username = "admin'; #'";
        password = "mikasa";

        String sql = "SELECT * FROM admin WHERE username = '" + username + "' AND password = '" + password + "'";

        System.out.println("SQL:\n" + sql);
        ResultSet resultSet = statement.executeQuery(sql);

        if(resultSet.next()){
            System.out.println("Success login: " + resultSet.getString("username"));
        } else {
            System.out.println("Unsuccessful login");
        }

        resultSet.close();
        statement.close();
        connection.close();
    }
}
