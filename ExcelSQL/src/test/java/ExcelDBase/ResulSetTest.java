package ExcelDBase;

import org.junit.jupiter.api.Test;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;

public class ResulSetTest {
    @Test
    void testExecuteQuery() throws SQLException {
        Connection connection = ConnectionUtil.getDataSource().getConnection();
        Statement statement = connection.createStatement();

        String sql = """
                SELECT * FROM customer;
                """;

        ResultSet resultSet = statement.executeQuery(sql);

        System.out.println("\nName  Email  Id");
        while (resultSet.next()){
            String id, name, email;
            name = resultSet.getString("name");
            email = resultSet.getString("email");
            id = resultSet.getString("id");

            System.out.println(
                    String.join("  ", name, email, id)
            );
        }

        resultSet.close();
        statement.close();
        connection.close();
    }
}
