package launcher;

import action.MyConfig;
import datatable.readDatatable;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import report.ReportWP;
import testcase.executeTest;

import java.io.File;

public class launch {
    public static void main(String[] args) throws Exception {
        String strDataTablePath, strFile, strExecutePath, strFileExecute, strFileExcel;
        String strChromeDriver, strTypeWp;
        String intStartData, intEndData, strDatatableSheetName,strColumnAction,strColumnStatus,strColumnKeterangan;


//        HashMap hasListDescription = new HashMap<Integer,String>();
//        ReportWP reportWP = new ReportWP();
//        reportWP.CreateWPExcel("C:\\Screens\\BTRD\\Running18052020\\Running18052020\\", "Running18052020", "HORIZONTAL_DYNAMIC");

        //Nama Sheet Datatable Utama
        MyConfig.strDatatableSheetName = "Datatable";
        strFileExecute = "ExecuteTable.xlsx";
        strDataTablePath =  System.getProperty("user.dir") + "\\DataTable\\";

        File folder = new File(strDataTablePath);
        File[] listOfFiles = folder.listFiles();

        //Get Value Table
        readDatatable objReadDatatable = new readDatatable();
        Sheet sheetTestcase = objReadDatatable.readTestcase(strDataTablePath, strFileExecute);

        //looping testcase from datatable
        int rowCount = sheetTestcase.getLastRowNum() - sheetTestcase.getFirstRowNum();
        for(int intCounterCaseFile=1; intCounterCaseFile<rowCount+1; intCounterCaseFile++) {
            Row row = sheetTestcase.getRow(intCounterCaseFile);
            if (row.getCell(1) != null) {
                if(!row.getCell(1).getStringCellValue().equalsIgnoreCase("")){
                    strFileExcel =  row.getCell(1).toString() + ".xlsx";
                    strChromeDriver = sheetTestcase.getRow(1).getCell(2).toString();
                    strTypeWp = sheetTestcase.getRow(1).getCell(3).toString();
                    strExecutePath = sheetTestcase.getRow(1).getCell(4).toString();
                    intStartData = sheetTestcase.getRow(1).getCell(5).toString();
                    intEndData = sheetTestcase.getRow(1).getCell(6).toString();// kalau null gmn

                    String[] paramExcecuteConfig = new String[10];

                    paramExcecuteConfig[0] = strChromeDriver;
                    paramExcecuteConfig[1] = strTypeWp;
                    paramExcecuteConfig[2] = strExecutePath;
                    paramExcecuteConfig[3] = intStartData ;
                    paramExcecuteConfig[4] = intEndData;
                    paramExcecuteConfig[5] = strDataTablePath;
                    paramExcecuteConfig[6] = ""; //Untuk isi nama file testCase didalam for
                    paramExcecuteConfig[7] = intCounterCaseFile+"";

                    //Create Folder
                    File file = new File(strExecutePath);
                    if (!file.exists()) {
                        file.mkdirs();
                        System.out.println("Directory is created!");
                    }
                    for (Integer i = 0; i < listOfFiles.length; i++){
                        strFile = listOfFiles[i].getName();
                        paramExcecuteConfig[6] = strFile;
                        if(strFileExcel.toString().toUpperCase().equals(strFile.toString().toUpperCase())){
                            executeTest objExecuteTest = new executeTest();
                            objExecuteTest.execute(paramExcecuteConfig);
                        }
                    }
                }
            }
        }
        System.exit(1);
    }



}
