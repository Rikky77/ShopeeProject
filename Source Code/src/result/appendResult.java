package result;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class appendResult {

    public static void appendCell(String[] paramToAppend, String resultFile) throws IOException, InterruptedException {

        File file = new File(resultFile);
        if(!file.exists()) {
            createResult.create(resultFile);
        }
        //Get the workbook instance for XLSX file
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file));
        //get spreadsheet
        XSSFSheet spreadsheet = workbook.getSheetAt(0);

          /*
        param0: Status
        param1: TestCase
        param2: Expected
        param3: Actual
        */
        String strStatus, strTestCase, strExpected, strActual, srtDataKe;
        strStatus = paramToAppend[0];
        strTestCase = paramToAppend[1];
        strExpected = paramToAppend[2];
        strActual = paramToAppend[3];
        srtDataKe = paramToAppend[4];

        int intLastRow = spreadsheet.getLastRowNum() + 1;
        if ( spreadsheet.getRow(intLastRow)==null){
            spreadsheet.createRow(intLastRow);
        }

        spreadsheet.getRow(intLastRow).createCell(0).setCellValue(srtDataKe);
        spreadsheet.getRow(intLastRow).createCell(1).setCellValue(strTestCase);
        spreadsheet.getRow(intLastRow).createCell(2).setCellValue(strExpected);
        spreadsheet.getRow(intLastRow).createCell(3).setCellValue(strActual);
        spreadsheet.getRow(intLastRow).createCell(4).setCellValue(strStatus);

        FileOutputStream outputStream = new FileOutputStream(new File(resultFile));
        workbook.write(outputStream);
        outputStream.close();



    }
}
