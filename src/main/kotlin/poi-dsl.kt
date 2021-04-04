import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream


fun main() {
    val workbook = xlsx {
        sheet("Java Books") {
            val bookData = arrayOf(
                arrayOf("Head First Java", "Kathy Serria", 79),
                arrayOf("Effective Java", "Joshua Bloch", 36),
                arrayOf("Clean Code", "Robert martin", 42),
                arrayOf("Thinking in Java", "Bruce Eckel", 35)
            )

            for ((rowCount, aBook) in bookData.withIndex()) {
                row(rowCount + 1) {
                    for ((columnCount, field) in aBook.withIndex()) {
                        cell(columnCount + 1) {
                            if (field is String) {
                                setCellValue(field)
                            } else if (field is Int) {
                                setCellValue(field.toDouble())
                            }
                        }
                    }
                }
            }
        }

        sheet("2nd sheet") {}
    }

    FileOutputStream("JavaBooks.xlsx").use { outputStream -> workbook.write(outputStream) }
}

fun xlsx(init: XSSFWorkbook.() -> Unit): XSSFWorkbook {
    val xlsx = XSSFWorkbook()
    xlsx.init()
    return xlsx
}

fun XSSFWorkbook.sheet(sheetname: String, init: XSSFSheet.() -> Unit): XSSFSheet {
    val sheet = createSheet(sheetname)
    sheet.init()
    return sheet
}

fun XSSFSheet.row(rownum: Int, init: XSSFRow.() -> Unit): XSSFRow {
    val row = createRow(rownum)
    row.init()
    return row
}

fun XSSFRow.cell(columnIndex: Int, init: XSSFCell.() -> Unit): XSSFCell {
    val cell = createCell(columnIndex)
    cell.init()
    return cell
}

