package sky.xls2csv
import org.apache.poi.hssf.usermodel._
import org.apache.poi.ss.usermodel._
import java.io._
import scala.collection.JavaConversions._
object xlswork {
	implicit def toFileInputStream(s: String): FileInputStream ={
					new FileInputStream(new File(s))
	}                                         //> toFileInputStream: (s: String)java.io.FileInputStream
	implicit val workbook = new HSSFWorkbook("OBB_OE_2012_05_W1_P01_C1.xls");
                                                  //> workbook  : org.apache.poi.hssf.usermodel.HSSFWorkbook = org.apache.poi.hssf
                                                  //| .usermodel.HSSFWorkbook@2914f7c0
	val numberOfSheet = workbook.getNumberOfSheets();
                                                  //> numberOfSheet  : Int = 5
	
	val sheet = workbook.getSheetAt(0)        //> sheet  : org.apache.poi.hssf.usermodel.HSSFSheet = org.apache.poi.hssf.userm
                                                  //| odel.HSSFSheet@1f957352
	
	(0 to workbook.getNumberOfSheets - 1).map(workbook.getSheetAt(_)).toList.size
                                                  //> res0: Int = 5
                                                  
  implicit def noHiddenSheet(i: Int ):Boolean = !(workbook.isSheetHidden(i) || workbook.isSheetVeryHidden(i))
                                                  //> noHiddenSheet: (i: Int)Boolean
 	def sheetAt(index: Int,filter: Int => Boolean = noHiddenSheet): Option[HSSFSheet] ={
 				if(filter(index))Some(workbook.getSheetAt(index))
 				else None
 	}                                         //> sheetAt: (index: Int, filter: Int => Boolean)Option[org.apache.poi.hssf.user
                                                  //| model.HSSFSheet]
 	
 	sheetAt(2)                                //> res1: Option[org.apache.poi.hssf.usermodel.HSSFSheet] = Some(org.apache.poi.
                                                  //| hssf.usermodel.HSSFSheet@180bed71)
 	
 	val sheets = (0 to workbook.getNumberOfSheets - 1).map(sheetAt(_)).toList.filter(_ != None)
                                                  //> sheets  : List[Option[org.apache.poi.hssf.usermodel.HSSFSheet]] = List(Some(
                                                  //| org.apache.poi.hssf.usermodel.HSSFSheet@180bed71), Some(org.apache.poi.hssf.
                                                  //| usermodel.HSSFSheet@26d82faa), Some(org.apache.poi.hssf.usermodel.HSSFSheet@
                                                  //| 3b2155d2))
  val s = sheet.iterator().toStream               //> s  : scala.collection.immutable.Stream[org.apache.poi.ss.usermodel.Row] = St
                                                  //| ream(org.apache.poi.hssf.usermodel.HSSFRow@2f101fa5, ?)
     s.size                                       //> res2: Int = 35
  val c  = s.tail.head.cellIterator().toStream    //> c  : scala.collection.immutable.Stream[org.apache.poi.ss.usermodel.Cell] = 
                                                  //| Stream(1.0, ?)
           c.size                                 //> res3: Int = 8
	
	def getValue(c: Cell): AnyVal = {
			c.getCellType() match{
                case Cell.CELL_TYPE_BOOLEAN => c.getBooleanCellValue()
                case Cell.CELL_TYPE_NUMERIC=> c.getNumericCellValue()
                case Cell.CELL_TYPE_STRING => c.getStringCellValue()
                case Cell.CELL_TYPE_BLANK =>  ""
                case _ => c.toString
              }
      }                                           //> getValue: (c: org.apache.poi.ss.usermodel.Cell)AnyVal
  s foreach {x =>
     val l = (for(v <- x.cellIterator) yield getValue(v))
     println(l mkString("",",","\n"))             //> 9.0
                                                  //| 
                                                  //| 1.0,5.0,1.0,10.0,1.0,0.0,2.0,32.0
                                                  //| 
                                                  //| 24.0,54.0,9.0,11.0,13.0,11.0,20.0,2.0
                                                  //| 
                                                  //| 1.0,4P39UQZXTH37RZIHFPA7Q66FK,true,true,SAPBEXq0001,3.0,4.0,false,true,fals
                                                  //| e,true,false,2.0,5.0,true,false,false,false,4.0,0FUND__ZVOTE,Vote,X,,,0001,
                                                  //| ,,,,,,,0,00,,,,00000000,K,0FUND__ZVOTE,A,,,H,,,,,,00000000,0000,,0,,,,,1,1,
                                                  //| 1,0,0,,,2,,0,4P39UPI64QV95LQCAUTTRSFF4,,,,,,4.0,4P39UQCW8LA383W4Y736W0AA8,4
                                                  //| P39UQKKRJVSQQFL415J62900,Amaun,0001,,S,,L,X,4.0,0BUS_AREA,0LOGSYS,Source Sy
                                                  //| stem,0000,X,,0,0,,,,4.0,4P39UQKKRJVSQQFL415J62900,0000000000000000000000000
                                                  //| ,,,,,,00,X,4,0,,,4.0,ZFISYR,1,P,I,EQ,2012,2012,,,,0,,0,,0,,20,,,0FISCYEAR,4
                                                  //| .0,0000000100,
                                                  //| 
                                                  //| 4.0,4P39UQCW8LA383W4Y736W0AA8,Structure,,X,X,0001,,Amaun,,,,,U,,00,,,,,K,4P
                                                  //| 39UQCW8LA383W4Y736W0AA8,A,,,,,,,,,00000000,0000,X,0,,,,,0,0,0,0,0,,,,,,4P39
                                                  //| UQCW8LA383W4Y736W0AA8,,,,,,4.0,0BUS_AREA,0SOURSYSTEM,Source system ID,0000,
                                                  //| X,,0,0,,,,4.0,PERDRANG,1,P,I,EQ,005,5,,,May,20
                                                  //| Output exceeds cutoff limit.
     }
	//c foreach {x => println(getValue(x))}
  print(1)                                        //> 1
	implicit def noHiddenSheet2(i: Int, w:Workbook ):Boolean = !(w.isSheetHidden(i) || w.isSheetVeryHidden(i))
                                                  //> noHiddenSheet2: (i: Int, w: org.apache.poi.ss.usermodel.Workbook)Boolean
	def sheetAt2(w:Workbook, index: Int,filter: (Int, Workbook) => Boolean = noHiddenSheet2): Option[HSSFSheet] ={
 				if(filter(index, workbook))Some(workbook.getSheetAt(index))
 				else None
 	}                                         //> sheetAt2: (w: org.apache.poi.ss.usermodel.Workbook, index: Int, filter: (In
                                                  //| t, org.apache.poi.ss.usermodel.Workbook) => Boolean)Option[org.apache.poi.h
                                                  //| ssf.usermodel.HSSFSheet]
	sheetAt2(workbook, 2)                     //> res4: Option[org.apache.poi.hssf.usermodel.HSSFSheet] = Some(org.apache.poi
                                                  //| .hssf.usermodel.HSSFSheet@180bed71)
                                                  
	
}