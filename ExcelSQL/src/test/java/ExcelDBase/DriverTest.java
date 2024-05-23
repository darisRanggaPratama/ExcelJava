package ExcelDBase;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.sql.Driver;
import java.sql.DriverManager;

public class DriverTest {
    @Test
    void testRegister(){
        try {
            Driver mysqlDriver = new com.mysql.cj.jdbc.Driver();
            DriverManager.registerDriver(mysqlDriver);
        }catch (Exception e){
            System.out.println("Fail :\n");
            Assertions.fail(e);
        }
    }
}
