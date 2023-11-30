package kotlins

import basic.Data
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException

class NullCellK {
    companion object {
        private val logger = System.getLogger("NullCell")
        fun checkNull() {
            try {
                // Get Excel file path
                val name = Data.fileXl
                val path = "./src/main/java/trial/"
                val excelFilePath = path + name

                // Open file
                val input = FileInputStream(excelFilePath)
                val workbook: Workbook = XSSFWorkbook(input)

                // Get worksheet (GAJI)
                val sheetName = Data.sheetXl
                val sheet = workbook.getSheet(sheetName)

                // Create new sheet
                val sheetNew = workbook.createSheet("Null")

                // Get Negative. Example: A2 to E46
                val firstRow = Data.beginRow
                val lastRow = Data.endRow
                val firstCol = Data.firstColumn
                val lastCol = Data.lastColumn

                // Create Cell Title
                val rowTitle = sheetNew.createRow(0)
                val titleNo = rowTitle.createCell(0)
                titleNo.setCellValue("No")
                val cellTitle = rowTitle.createCell(1)
                cellTitle.setCellValue("Cell Null/Empty")
                println("\nNo  Null/Empty")

                // Check null/empty/missing/blank cell
                var rowIndex = 1
                var no = 0

                for (rowIdx in firstRow until lastRow) {
                    val row = sheet.getRow(rowIdx)

                    for (colIdx in firstCol until lastCol) {
                        val cell = row?.getCell(colIdx)

                        // Check null/empty/missing/blank value
                        if (cell == null ||
                            (cell.cellType == CellType.STRING && cell.stringCellValue.trim().isEmpty())
                        ) {
                            no++

                            // Write blank address cell in new sheet
                            val newRow = sheetNew.createRow(rowIndex++)
                            val newNo = newRow.createCell(0)
                            val newCellAddress = newRow.createCell(1)
                            val cellAddress = CellReference.convertNumToColString(colIdx) + (rowIdx + 1)
                            newNo.setCellValue(no.toDouble())
                            newCellAddress.setCellValue(cellAddress)

                            // Write in console
                            println("$no   $cellAddress")
                        }

                    }

                }

                // Save file
                val output = FileOutputStream("./src/main/java/output/Null.xlsx")
                workbook.write(output)

                // Close Excel file
                workbook.close()
                input.close()
                output.close()
            } catch (e: IOException) {
                val noFile = "\nFile not found: \n$e"
                logger.log(System.Logger.Level.ERROR, noFile)
            } catch (e: NullPointerException) {
                val blank = "\nBlank Cell: \n$e"
                logger.log(System.Logger.Level.ERROR, blank)
            }
        }
    }
}

