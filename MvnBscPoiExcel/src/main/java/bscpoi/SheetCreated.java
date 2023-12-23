package bscpoi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class SheetCreated {
    public static void main(String[] args) throws Exception {
        // Create blank workbook
        XSSFWorkbook book = new XSSFWorkbook();

        // Create blank worksheet
        XSSFSheet sheet = book.createSheet("data");

        // Create row object
        XSSFRow row;

        // Data
        Map<String, Object[]> info = new TreeMap<String, Object[]>();

        info.put("1", new Object[]{
                "ID", "Name", "Position"
                });

        info.put("2", new Object[]{
                "id1", "Eren", "Mobile Developer"
        });

        info.put("3", new Object[]{
                "id2", "Mikasa", "Desktop Developer"
        });

        info.put("4", new Object[]{
                "id3", "Armin", "Web Developer"
        });

        info.put("5", new Object[]{
                "id4", "Levi", "Backend Developer"
        });

        info.put("6", new Object[]{
                "id5", "Erwin", "Project Manager"
        });

        // Iterate over data & write to sheet
        Set<String> keyid = info.keySet();
        int rowid = 0;

        for (String key: keyid){
            row = sheet.createRow(rowid++);
            Object[] arrayObj = info.get(key);
            int cellid = 0;
            for (Object obj: arrayObj){
                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String) obj);
            }
        }

        // Write data to workbook
        FileOutputStream outFile = new FileOutputStream(
                new File("./src/main/java/output/newSheet.xlsx")
        );

        book.write(outFile);
        outFile.close();
        System.out.println("Create new workbook");
    }
}
