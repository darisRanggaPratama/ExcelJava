package SqlExcel;

import SqlExcel.LearnSQL.ExportData;
import SqlExcel.LearnSQL.ImportData;

import java.io.IOException;
import java.sql.SQLException;

public class App
{
    public static void main( String[] args ) throws SQLException, IOException {
       // ExportData.CreateExcel();
        ImportData.readExcel();
    }
}
