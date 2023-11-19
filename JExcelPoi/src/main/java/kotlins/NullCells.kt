package kotlins

import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException

fun main() {
    try {
        // File path
        val filePath = "./src/main/java/trial/sample.xlsx"

        // Open file
        val input = FileInputStream(filePath)
        val workbook = XSSFWorkbook(input)

        // Get sheet name "GAJI"
        val sheet = workbook.getSheet("GAJI")

        // Create new sheet
        val sheetNew = workbook.createSheet("null")

        // Create cell title
        val rowTitle = sheetNew.createRow(0)
        var noTitle = rowTitle.createCell(0)
        val cellTitle = rowTitle.createCell(1)
        noTitle.setCellValue("No")
        cellTitle.setCellValue("Alamat Cell Null/Empty")
        println("\nNull/Empty")

        // Check null/empty/missing/blank cell
        var rowIndex = 1
        var no = 0;
        for (rowIdx in 1 until 47) {
            val row = sheet.getRow(rowIdx)

            for (colIdx in 0 until 6) {
                val cell = row?.getCell(colIdx)

                // Check null/empty/missing/blank value
                if (cell == null || (cell.cellType == CellType.STRING && cell.stringCellValue.trim().isEmpty())) {
                    // Write blank address cell in new sheet
                    val nullRow = sheetNew.createRow(rowIndex++)
                    val cellAddress = nullRow.createCell(0)
                    no++
                    cellAddress.setCellValue("$no  ${CellReference.convertNumToColString(colIdx)}${rowIdx + 1}")

                    println("$no  ${CellReference.convertNumToColString(colIdx)}${rowIdx + 1}")
                }
            }
        }

        // Save file
        val output = FileOutputStream("./src/main/java/output/KNulls.xlsx")
        workbook.write(output)
        workbook.close()
        input.close()
        output.close()
    } catch (e: IOException) {
        e.printStackTrace()
    }
}

