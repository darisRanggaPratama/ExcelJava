package SqlExcel;

import SqlExcel.LearnSQL.ExportCSV;
import SqlExcel.LearnSQL.ExportData;
import SqlExcel.LearnSQL.ImportCSV;
import SqlExcel.LearnSQL.ImportData;

import java.io.IOException;
import java.sql.SQLException;

public class App
{
    public static void main( String[] args ) throws SQLException, IOException {
        ImportData.readExcel();
        ImportCSV.readCSV();
         ExportData.CreateExcel();
        ExportCSV.createCSV();
    }
}
