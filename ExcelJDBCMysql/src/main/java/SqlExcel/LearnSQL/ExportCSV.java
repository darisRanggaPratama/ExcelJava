package SqlExcel.LearnSQL;

import SqlExcel.DatabaseConnect;

import java.io.FileWriter;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;


public class ExportCSV {
    static Connection connection = null;
    static String sql, file;
    static PreparedStatement preState = null;
    static ResultSet resultSet = null;
    static int columnCount;
    static FileWriter fileWriter = null;

    public static void createCSV() throws SQLException, IOException {
        // Connect to database
        connection = DatabaseConnect.getDataSource().getConnection();

        sql = """
                SELECT * FROM customer;
                """;

        preState = connection.prepareStatement(sql);
        resultSet = preState.executeQuery();
        // Get matadata to know column name
        columnCount = resultSet.getMetaData().getColumnCount();
        file = "./src/main/java/created/EmailCSV.csv";
        fileWriter = new FileWriter(file);

        // Write Header/title row
        for (int i = 1; i <= columnCount; i++){
            fileWriter.append(resultSet.getMetaData().getColumnName(i));
            if (i < columnCount){
                fileWriter.append(";");
            }
        }
        fileWriter.append("\n");

        // Write data row
        while (resultSet.next()){
            for (int i = 1; i <= columnCount; i++){
                fileWriter.append(resultSet.getString(i));
                if (i < columnCount){
                    fileWriter.append(";");
                }
            }
            fileWriter.append("\n");
        }

        System.out.println("\nCreate CSV file\n");

        try{
            // Close resource
            if (resultSet != null){
                resultSet.close();
            }
            if (preState != null){
                preState.close();
            }
            if (connection != null){
                connection.close();
            }
            if (fileWriter != null){
                fileWriter.flush();
            }
            if (fileWriter != null){
                fileWriter.close();
            }
        } catch (SQLException | IOException e){
            System.out.println("\nFail:\n" + e);
        }

    }
}
