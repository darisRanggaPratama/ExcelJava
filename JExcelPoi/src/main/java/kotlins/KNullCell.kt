package kotlins

import basic.Data
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFClientAnchor
import org.apache.poi.xssf.usermodel.XSSFRichTextString
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException
import java.util.*


class KNullCell {
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
                var rowIdx = 0
                var colIdx = 0
                var no = 0
                workbook.missingCellPolicy = Row.MissingCellPolicy.CREATE_NULL_AS_BLANK
                rowIdx = firstRow
                for (rowIdx in rowIdx..lastRow step 1) {
                    val idxRow = Optional.ofNullable(rowIdx)
                        .orElse(0)
                    val row = sheet.getRow(idxRow)
                    colIdx = firstCol
                    for (colIdx in colIdx..lastCol step 1) {
                        val idxCol = Optional.ofNullable(colIdx)
                            .orElse(0)
                        val cell = row.getCell(idxCol)

                        // Check null/empty/missing/blank value
                        if (cell == null || cell.cellType == CellType.BLANK || cell.cellType == CellType.STRING && cell.stringCellValue.trim { it <= ' ' }
                                .isBlank() || cell.cellType == CellType.STRING && cell.stringCellValue.trim { it <= ' ' }
                                .isEmpty()) {
                            no++

                            // Add comment
                            val comment = sheet.createDrawingPatriarch().createCellComment(
                                XSSFClientAnchor(
                                    0, 0, 0, 0,
                                    colIdx.toShort().toInt(), rowIdx, (colIdx + 1).toShort().toInt(), rowIdx + 1
                                )
                            )
                            comment.string = XSSFRichTextString("Null")
                            cell!!.cellComment = comment

                            // Add Background color
                            val style = workbook.createCellStyle()
                            style.fillForegroundColor = IndexedColors.LIGHT_ORANGE.getIndex()
                            style.fillPattern = FillPatternType.SOLID_FOREGROUND
                            cell.cellStyle = style

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
