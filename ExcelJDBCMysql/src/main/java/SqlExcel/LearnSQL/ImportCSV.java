package SqlExcel.LearnSQL;

import SqlExcel.DatabaseConnect;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;

public class ImportCSV {
    static Connection connect;
    static FileReader readFile;
    static BufferedReader buffRead;
    static String filePath, sql, text = null, id, email, name;
    static String[] data;
    static int count = 0, batchSize = 20;
    static PreparedStatement preState;

    public static void readCSV() {
        try {
            connect = DatabaseConnect.getDataSource().getConnection();
            filePath = "./src/main/java/created/importCSV.csv";
            readFile = new FileReader(filePath);
            connect.setAutoCommit(false);

            sql = """
                    INSERT INTO customer(id, name, email) VALUES(?, ?, ?);
                    """;

            preState = connect.prepareStatement(sql);
            buffRead = new BufferedReader(readFile);
            buffRead.readLine();

            while ((text = buffRead.readLine()) != null) {
                data = text.split(";");
                id = data[0];
                name = data[1];
                email = data[2];

                preState.setString(1, id);
                preState.setString(2, name);
                preState.setString(3, email);
                preState.addBatch();

                if (count % batchSize == 0) {
                    preState.executeBatch();
                }
            }
            buffRead.close();
            // Execute the remaining queries
            preState.executeBatch();
            connect.commit();
            connect.close();
            System.out.println("\nInsert data successfully \n");

        } catch (IOException | SQLException e) {
            System.out.println("\nFail to insert data: \n" + (e));
        }
    }
}
