package ExcelDBase;

import org.junit.jupiter.api.Test;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;

public class PreparedStatementTest {
    @Test
    void testPreparedStatement() throws SQLException {
        Connection connect = ConnectionUtil.getDataSource().getConnection();
        String username = "admin", password = "armin";

        String sql = "SELECT * FROM admin WHERE username = ? AND password = ?";

        PreparedStatement preState = connect.prepareStatement(sql);
        preState.setString(1, username);
        preState.setString(2, password);
        ResultSet resultSet =  preState.executeQuery();

        if(resultSet.next()){
            System.out.println("Success login: " + resultSet.getString("username"));
        } else {
            System.out.println("Unsuccessful login");
        }

        preState.close();
        connect.close();
    }
}
