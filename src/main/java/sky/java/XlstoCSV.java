package sky.java;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

class XlstoCSV {

    static void xls(File inputFile, File outputFile) {
        // For storing data into CSV files
        StringBuffer data = new StringBuffer();
        try {
            FileOutputStream fos = new FileOutputStream(outputFile);

            // Get the workbook object for XLS file
            HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(inputFile));
            int numberOfSheet = workbook.getNumberOfSheets();
            
            for (int i = 0; i < numberOfSheet; i++) {
                if (!workbook.isSheetHidden(i) && !workbook.isSheetVeryHidden(i)) {
                    HSSFSheet sheet = workbook.getSheetAt(i);
                    Cell cell;
                    Row row;

                    // Iterate through each rows from the sheet
                    Iterator<Row> rowIterator = sheet.iterator();
                    while (rowIterator.hasNext()) {
                        row = rowIterator.next();
                        // For each row, iterate through each columns
                        Iterator<Cell> cellIterator = row.cellIterator();
                        while (cellIterator.hasNext()) {
                            cell = cellIterator.next();

                            switch (cell.getCellType()) {
                                case Cell.CELL_TYPE_BOOLEAN:
                                    data.append(cell.getBooleanCellValue());
                                    break;

                                case Cell.CELL_TYPE_NUMERIC:
                                    data.append(cell.getNumericCellValue());
                                    break;

                                case Cell.CELL_TYPE_STRING:
                                    data.append(cell.getStringCellValue());
                                    break;

                                case Cell.CELL_TYPE_BLANK:
                                    data.append("");
                                    break;

                                default:
                                    data.append(cell);
                            }
                            if (cellIterator.hasNext()) {
                                data.append(",");
                            }
                        }
                        data.append("\n");
                    }
                }
            }
            fos.write(data.toString().getBytes());
            fos.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) {
        File inputFile = new File("src/main/resources/test.xls");
        File outputFile = new File("src/main/resources/output.csv");

        xls(inputFile, outputFile);
    }
}
