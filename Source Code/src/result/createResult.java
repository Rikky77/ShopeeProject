package result;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class createResult {
    public static void create(String resultFile) throws IOException {

        //Create Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Create a blank spreadsheet
        XSSFSheet spreadsheet = workbook.createSheet("Report");

        //Create row object
        XSSFRow row;

        //This data needs to be written (Object[])
        Map< String, Object[] > empinfo = new TreeMap< String, Object[] >();
        empinfo.put( "1", new Object[] {"No","TestCase", "Expected Result", "Actual Result", "Status"});

        //Iterate over data and write to sheet
        Set< String > keyid = empinfo.keySet();
        int rowid = 0;
        for (String key : keyid){
            row = spreadsheet.createRow(rowid++);
            Object [] objectArr = empinfo.get(key);
            int cellid = 0;
            for (Object obj : objectArr){
                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String)obj);
            }
        }

        FileOutputStream out = new FileOutputStream(new File(resultFile));
        //write operation workbook using file out object
        workbook.write(out);
        out.close();
        System.out.println(" createworkbook.xlsx written successfully");
    }
}
