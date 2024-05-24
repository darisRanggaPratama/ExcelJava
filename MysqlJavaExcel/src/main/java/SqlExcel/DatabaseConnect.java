package SqlExcel;

import com.zaxxer.hikari.HikariConfig;
import com.zaxxer.hikari.HikariDataSource;

public class DatabaseConnect {
    private static HikariDataSource dataSource;
    static{
        // Connection Pool (CP)
        HikariConfig config = new HikariConfig();
        config.setDriverClassName("com.mysql.cj.jdbc.Driver");
        config.setJdbcUrl("jdbc:mysql://localhost:3306/java_database?serverTimezone=Asia/Jakarta");
        config.setUsername("rangga");
        config.setPassword("rangga");

        // Configure: CP
        config.setMaximumPoolSize(10);
        config.setMinimumIdle(5);
        config.setIdleTimeout(60_000);
        config.setMaxLifetime(10 * 60_000);

        dataSource = new HikariDataSource(config);
    }
    public static HikariDataSource getDataSource(){
        return dataSource;
    }
}
