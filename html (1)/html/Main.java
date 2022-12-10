import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.lang.*;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main
{
    public static void main(String[]args) {
        File inputFile = new File("C:/Users/Rie/Pictures/cat201/test.xlsx");
        File outputFile = new File("C:/Users/Rie/Pictures/cat201/test.CSV");
        StringBuilder data = new StringBuilder();
        try {
            FileInputStream fis = new FileInputStream(inputFile);
            Workbook wb;
            if (inputFile.getName().endsWith(".xlsx")) {
                wb= new XSSFWorkbook(fis);
            } else if (inputFile.getName().endsWith(".xls")) {
                wb = new HSSFWorkbook(fis);
            } else {
                throw new Exception("File not supported!");
            }
            Sheet sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            if (rowIterator.hasNext()) {
                do {
                    Row row = rowIterator.next();
                    Iterator<Cell> cellIterator = row.cellIterator();
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        switch (cell.getCellType()) {
                            case BOOLEAN:
                                data.append(cell.getBooleanCellValue()).append(",");
                                break;
                            case NUMERIC:
                                data.append(cell.getNumericCellValue()).append(",");
                                break;
                            case STRING:
                                data.append(cell.getStringCellValue()).append(",");
                                break;
                            case BLANK:
                                data.append(cell).append(",");
                            default:
                                data.append(cell).append(",");
                        }
                    }
                    data.append('\n');
                } while (rowIterator.hasNext());
            }
            FileOutputStream fos = new FileOutputStream(outputFile);
            fos.write(data.toString().getBytes());
            fos.close();

        } catch (Exception e)
          {
              e.printStackTrace();
          }
        System.out.println("Conversion of an Excel file to CSV file is done!");
    }


}
