package excelMaven;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class CreateSpreadsheet {
    static System.Logger logger = System.getLogger("CheckWorkbook");
    static File file;
    static FileOutputStream result;
    static XSSFWorkbook workbook;
    static XSSFSheet worksheet;
    static XSSFRow row;
    static String txtInfo;

    public static void createWorksheet() {
        try {
            workbook = new XSSFWorkbook();
            worksheet = workbook.createSheet("Person Data");
            Map<String, Object[]> info = new TreeMap<>();

            info.put("1", new Object[]{"ID", "Name", "Job", "Age"});
            info.put("2", new Object[]{"01", "Eren", "HRD", "23"});
            info.put("3", new Object[]{"02", "Mikasa", "Accounting", "24"});
            info.put("4", new Object[]{"03", "Armin", "Engineering", "25"});
            info.put("5", new Object[]{"04", "Erwin", "Finance", "35"});
            info.put("6", new Object[]{"05", "Levi", "Procurement", "36"});

            Set<String> keyID = info.keySet();
            int rowID = 0, cellID;

            for (String key : keyID) {
                row = worksheet.createRow(rowID++);
                Object[] objArray = info.get(key);

                cellID = 0;
                for (Object obj : objArray) {
                    Cell cell = row.createCell(cellID++);
                    cell.setCellValue((String) obj);
                }
            }

            file = new File("./src/main/java/checking/SheetCreated.xlsx");
            result = new FileOutputStream(file);

            workbook.write(result);
            result.close();

            txtInfo = "File created successfully";
            logger.log(System.Logger.Level.ERROR, txtInfo);

        } catch (Exception e) {
            txtInfo = "File not found: \n" + e;
            logger.log(System.Logger.Level.ERROR, txtInfo);
        }
    }
}
