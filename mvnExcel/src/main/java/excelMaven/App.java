package excelMaven;

import java.lang.System.Logger;

public class App 
{
    static Logger logger = System.getLogger("BlankWorkbook");
    public static void main( String[] args ) {

        try{
           // CreateBlankWorkbook.blankWorkbook();
           // CheckExistFile.openExistWorkbook();
            // CreateSpreadsheet.createWorksheet();
           // ReadSpreadsheet.readWorksheet();
          //  DataTypes.ExcelData();
            CellStyles.styledCell();
        } catch (Exception e){
            String noFile = "Fail: \n" + e;
            logger.log(System.Logger.Level.ERROR, noFile);
        }
    }
}
