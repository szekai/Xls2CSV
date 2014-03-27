/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

package sky.xls2csv
import java.io._
import org.apache.poi.hssf.usermodel._
import org.apache.poi.ss.usermodel._
import scala.collection.JavaConversions._

object Xls2csv {
  
  implicit def toFileInputStream(s: String) = new FileInputStream(new File(s))
  implicit def toFile(s:String) = new File(s)
  
  def xls(inputFile: String, outputFile: String): Unit = {
    // For storing data into CSV files
    val data = new StringBuffer();
    val fos = new FileOutputStream(outputFile);

    // Get the workbook object for XLS file
    val workbook = new HSSFWorkbook(inputFile)
    def notHiddenSheet(i: Int, w: Workbook): Boolean = !(w.isSheetHidden(i) || w.isSheetVeryHidden(i))

    def sheetAt(w: Workbook, index: Int, filter: (Int, Workbook) => Boolean = notHiddenSheet): Option[HSSFSheet] = 
      if (filter(index, workbook)) Some(workbook.getSheetAt(index)) else None

    val sheets = (0 to workbook.getNumberOfSheets - 1).map(sheetAt(workbook, _)).filter(_ != None)
    
    def getValue(c: Cell): AnyVal = {
      c.getCellType() match {
        case Cell.CELL_TYPE_BOOLEAN => c.getBooleanCellValue()
        case Cell.CELL_TYPE_NUMERIC => c.getNumericCellValue()
        case Cell.CELL_TYPE_STRING  => c.getStringCellValue()
        case Cell.CELL_TYPE_BLANK   => ""
        case _                      => c.toString
      }
    }

    sheets foreach {
      s =>
        // Iterate through each rows from the sheet
        s.iterator foreach {
          r =>
            // For each row, iterate through each columns
            r.iterator foreach {
              c =>
                {
                  val v = for (v <- c.cellIterator) yield getValue(v)
                  val vs = v mkString ("", ",", "\n")
                  fos.write(vs.getBytes())
                }
            }
        }
    }
    fos.close();
  }

  def main(args: Array[String]) {
    xls("src/main/resources/test.xls", "src/main/resources/output.csv")
  }
}
