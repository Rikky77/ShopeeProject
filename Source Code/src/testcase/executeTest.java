package testcase;

import action.MyConfig;
import action.keyword;
import datatable.readDatatable;
import log.appendLog;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import report.ReportWP;
import report.appendReport;

import java.awt.*;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.sql.Timestamp;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.TimeUnit;

public class executeTest {
    public static boolean bolValidationNextStep;
    public static Map<String, String> mapValueTemp = new HashMap<String, String>();
    public static HashMap hasListDescription = new HashMap<Integer,String>();
    public static HashMap hasCounterDataListDescription = new HashMap<Integer,HashMap>();

    public static int intWPCounter = 1;
    public static int intSpecificWPCounter = 0;
    public static String strTypeWp = "";
    public static String strNameApp = "";
    public static int intCounterRow = 1;

    //Untuk FOR/WHILE/DO WHILE
    public static int intBackToPointCounterRow = -1;
    public static int intCounterLooping = -1;
    public static int intTotalLooping = -1;
    public static int intStartDataLooping = -1;

    public void execute(String[] paramExcecuteConfig) throws Exception {
        Timestamp timestamp = new Timestamp(System.currentTimeMillis());
        //Global Variable for WP
        strTypeWp = paramExcecuteConfig[1];
        String datatableFile = paramExcecuteConfig[6];
        String logPath = paramExcecuteConfig[2];
        String chromeDriver = paramExcecuteConfig[0];
        String datatablePath =  paramExcecuteConfig[5];
        String strPathOutputWP = logPath  + datatableFile.replace(".xlsx","") + "\\WP-" + datatableFile;
        strNameApp = datatableFile.substring(0,4);

        //log, screenshot, report
        String strLogfile = logPath + "$" + datatableFile.replace(".xlsx", "") + timestamp.getTime() + ".txt";
        String strReportfile = logPath + "$" + datatableFile.replace(".xlsx", "") + timestamp.getTime() + ".docx";
        String strResultFile = logPath + "$" + datatableFile.replace("cl.xlsx", "") + timestamp.getTime() + ".xlsx";

        appendLog.append(strLogfile, "Automation test started");

        System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir") + "\\src\\externalDriver\\" + chromeDriver + ".exe");
        WebDriver webDriver = new ChromeDriver();
        webDriver.manage().window().maximize();
        webDriver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        appendLog.append(strLogfile, "Chrome browser launched successfully");

        keyword objKeyword = new keyword(webDriver);
        readDatatable objReadDatatable = new readDatatable();

        String strStartData = paramExcecuteConfig[3];
        int intStartData = 2;
        if(!strStartData.equals("")){
            intStartData = (Integer.parseInt(strStartData.replace(".0","")))+1; //karena di kodingan kevin j = 2
        }

        String strEndData =  paramExcecuteConfig[4];
        int intEndData = 2;
        if(!strEndData.equals("")){
            intEndData = Integer.parseInt(strEndData.replace(".0",""))+1;
        }


        int intRowTemp = 0;
        Row rowTemp = null;

        //For data table
        for (int intCounterTestCase = intStartData; intCounterTestCase <= intEndData; intCounterTestCase++) {
            hasListDescription = new HashMap<Integer,String>();
            hasCounterDataListDescription.put(intCounterTestCase-1,hasListDescription);
            intWPCounter = 1;
            intSpecificWPCounter = 0;

            bolValidationNextStep = true;

            //EDIT BY KEH
            //Specific TestCase Klo ada Action
            Workbook workbookTestcase;
            Sheet sheetTestcase;

            Sheet sheetDatatable = objReadDatatable.readDatatable(datatablePath, datatableFile);
            int intColumnAction =  GetCounterColumn(sheetDatatable,"ACTION");

            if (intColumnAction == -1){
                sheetTestcase = objReadDatatable.readTestcase(datatablePath, datatableFile);;
                workbookTestcase = sheetTestcase.getWorkbook();
            }else{
                Row rowAction = sheetDatatable.getRow(intCounterTestCase-1);
                sheetTestcase = objReadDatatable.readSpecificTestcase(datatablePath+datatableFile,rowAction.getCell(intColumnAction).getStringCellValue());
                workbookTestcase = sheetTestcase.getWorkbook();
            }

            int rowCount = sheetTestcase.getLastRowNum() - sheetTestcase.getFirstRowNum();

            //EDIT BY KEH

            //for testcase
            for (intCounterRow = 1; intCounterRow < rowCount + 1; intCounterRow++) {
                Row row = sheetTestcase.getRow(intCounterRow);
                if (bolValidationNextStep == false) {
                    if (row.getCell(0) != null) {
                        break;
                        // bolValidationNextStep = true;
                    }
                }
                if (bolValidationNextStep == true) {
                    if (row.getCell(0) != null && !row.getCell(0).getStringCellValue().equalsIgnoreCase("")) {    //new testcase
                        if (intRowTemp != 0) {
                         /*
                        param0: keyword
                        param1: pagename
                        param2: objectname
                        param3: objecttype
                        param4: value
                        param5: logfile
                        param6: reportfile
                        param7: resultFile
                        param8: testcase
                        */

                            String[] paramExpected = new String[11];

                            paramExpected[0] = rowTemp.getCell(6).toString();
                            paramExpected[1] = (rowTemp.getCell(7) == null) ? "" : rowTemp.getCell(7).toString();
                            paramExpected[2] = (rowTemp.getCell(8) == null) ? "" : rowTemp.getCell(8).toString();
                            paramExpected[3] = (rowTemp.getCell(9) == null) ? "" : rowTemp.getCell(9).toString();
                            paramExpected[4] = (rowTemp.getCell(10) == null) ? "" : rowTemp.getCell(10).toString();
                            paramExpected[5] = strLogfile;
                            paramExpected[6] = strReportfile;
                            paramExpected[7] = strResultFile;
                            paramExpected[8] = rowTemp.getCell(0).toString();
                            paramExpected[9] = datatablePath.toString() + datatableFile.toString();
                            paramExpected[10] = String.valueOf(intCounterTestCase-1);

                            objKeyword.expected(paramExpected, rowTemp);

                        }
                        intRowTemp = row.getRowNum();
                        rowTemp = row;

                        appendLog.append(strLogfile, "New testcase started ->" + row.getCell(0).toString());
                        appendReport.appendText(strReportfile, "Testcase: " + row.getCell(0).toString());
                    } else {

                        String[] param = new String[8];

                        param[0] = row.getCell(1).toString();
                        param[1] = (row.getCell(2) == null) ? "" : row.getCell(2).toString();
                        param[2] = (row.getCell(3) == null) ? "" : row.getCell(3).toString();
                        param[3] = (row.getCell(4) == null) ? "" : row.getCell(4).toString();
                        Cell cell = row.getCell(5);
//                        param[4] = (row.getCell(5) == null) ? "" : row.getCell(5).toString();

                        if (row.getCell(5) == null) {
                            param[4] = "";
                        } else {
                            param[4] =  ProcessFormulaCell(workbookTestcase,row,cell,intCounterTestCase);
                        }
                        param[5] = datatablePath.toString() + datatableFile.toString();
                        param[6] = datatableFile.replace(".xlsx","");

                        int intCounterData = intCounterTestCase-1;
                        objKeyword.perform(param, strLogfile, strReportfile, intCounterData);
                    }
                }

                if (intCounterRow == rowCount && intCounterTestCase == intEndData) {
                    String[] paramExpected = new String[11];

                    paramExpected[0] = rowTemp.getCell(6).toString();
                    paramExpected[1] = (rowTemp.getCell(7) == null) ? "" : rowTemp.getCell(7).toString();
                    paramExpected[2] = (rowTemp.getCell(8) == null) ? "" : rowTemp.getCell(8).toString();
                    paramExpected[3] = (rowTemp.getCell(9) == null) ? "" : rowTemp.getCell(9).toString();
                    paramExpected[4] = (rowTemp.getCell(10) == null) ? "" : rowTemp.getCell(10).toString();
                    paramExpected[5] = strLogfile;
                    paramExpected[6] = strReportfile;
                    paramExpected[7] = strResultFile;
                    paramExpected[8] = rowTemp.getCell(0).toString();
                    paramExpected[9] = datatablePath.toString() + datatableFile.toString();
                    paramExpected[10] = String.valueOf(intCounterTestCase-1);

                    objKeyword.expected(paramExpected, rowTemp);
                }
            }
            String[] paramUpdateStatus = new String[6];
            paramUpdateStatus[0] = strPathOutputWP;
            paramUpdateStatus[3] = String.valueOf(intCounterTestCase-1);
            paramUpdateStatus[4] = datatablePath.toString() ;
            paramUpdateStatus[5] = paramExcecuteConfig[7];

            if (bolValidationNextStep){
                paramUpdateStatus[1] = "PASSED";
                paramUpdateStatus[2] = "PASSED Perform -> Data ke-" + (intCounterTestCase-1) ;
                appendStatusExecuteTable(paramUpdateStatus);

            }else{
                paramUpdateStatus[1] = "FAILED";
                paramUpdateStatus[2] = "ERROR Perform -> Data ke-" + (intCounterTestCase-1) ;
                appendStatusExecuteTable(paramUpdateStatus);
              //  break;
            }


        }
//      Agar di WP biar tetap ada datatable
        File fleTestCaseSource = new File(datatablePath +  datatableFile);
        File fleTestCaseDestination = new File(strPathOutputWP);
        Files.copy(fleTestCaseSource.toPath(),fleTestCaseDestination.toPath(), StandardCopyOption.REPLACE_EXISTING);

        CreateWP(strLogfile,  datatableFile.replace(".xlsx",""));
        //sendEmail(webDriver, strReportfile);
        webDriver.close();
        webDriver.quit();
    }

    public void sendEmail(WebDriver webDriver, String strReportfile) throws InterruptedException {
        webDriver.get("https://jkt10.mail.bca.co.id/owa");
        webDriver.manage().timeouts().implicitlyWait(3,TimeUnit.SECONDS);
        webDriver.findElement(By.xpath("//*[@id=\"newmsgc\"]")).click();
        //Switch page
        String MainWindow = webDriver.getWindowHandle();
        Boolean flgPage = false;

        Set<String> s1 = webDriver.getWindowHandles();
        Iterator<String> i1 = s1.iterator();

        while (i1.hasNext())
        {
            String ChildWindow = i1.next();
            if(!MainWindow.equalsIgnoreCase(ChildWindow)){
                webDriver.switchTo().window(ChildWindow);
                flgPage = true;
            }
        }

        webDriver.findElement(By.xpath("//*[@id=\"divTo\"]")).sendKeys("IMAM ARIF FADHILAH; VERAWATY JAHJA; APRILIOAN YOPY SAPUTRA; RUDI BUDIYANTO");
        //webDriver.findElement(By.xpath("//*[@id=\"divTo\"]")).sendKeys("Dewi Indah Sari Sudjoko; RIKKY");
        webDriver.findElement(By.xpath("//*[@id=\"divCc\"]")).sendKeys("Dewi Indah Sari Sudjoko; RIKKY");


        webDriver.findElement(By.xpath("//*[@id=\"txtSubj\"]")).sendKeys("TRANSACTION SCHEDULE TRANSFER VIRTUAL ACCOUNT");
        webDriver.findElement(By.xpath("/html/body")).sendKeys("TRANSACTION SCHEDULE TRANSFER VIRTUAL ACCOUNT");

        webDriver.switchTo().frame("ifBdy");

        webDriver.findElement(By.xpath("/html/body")).sendKeys("Dear All, \n FYI \nTransaksi Transfer Virtual Account telah di proses, berikut No Rekening dan Nomor Virtual Account yang dijalankan : \n\n\n   Nomor Rekening: 0151100031 dan Nomor Virtual Account: 808888161964775\n  Nomor Rekening: 0151100031 dan Nomor Virtual Account: 808888161964775\n  Nomor Rekening: 0151100031 dan Nomor Virtual Account: 808888161964775\n  Nomor Rekening: 0151100031 dan Nomor Virtual Account: 808888161964775\n  Nomor Rekening: 0151100031 dan Nomor Virtual Account: 808888161964775\n  Nomor Rekening: 0151100031 dan Nomor Virtual Account: 808888161964775\n  Nomor Rekening: 0151100031 dan Nomor Virtual Account: 808888161964775\n  Nomor Rekening: 0151100031 dan Nomor Virtual Account: 808888161964775\n  Nomor Rekening: 0151100031 dan Nomor Virtual Account: 808888161964775\n  Nomor Rekening: 0151100031 dan Nomor Virtual Account: 808888161964775\n\nBerikut telah terlampir WP dari  Transfer Virtual Account:\n"+ strReportfile.replace("C:","\\"+"\\10.5.207.93"));
        Thread.sleep(1000);
        shortKey("ENTER");

        webDriver.switchTo().defaultContent();

        webDriver.findElement(By.xpath("//*[@id=\"send\"]")).click();

        webDriver.switchTo().window(MainWindow);
    }

    public void CreateWP(String strLogfile, String strFilename){
        ReportWP reportWP = new ReportWP();
        reportWP.CreateWPExcel(new File(strLogfile).getParent() +"\\" + strFilename+"\\", strFilename, executeTest.strTypeWp);

    }

    public static void inputKey(String data) {
        try {

            String[] arr = data.split("");
            for (Integer i = 0; i < arr.length; i++) {
                //panggil keyPress
                keyPress(arr[i]);
            }
        } catch (Exception e) {
            System.out.println("failed to perform inputKey");
        }
    }

    public static void keyPress(String data) {
        try {
            Robot r = new Robot();

            if (data.equals("a")) {
                r.keyPress(KeyEvent.VK_A);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_A);
            } else if (data.equals("b")) {
                r.keyPress(KeyEvent.VK_B);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_B);
            } else if (data.equals("c")) {
                r.keyPress(KeyEvent.VK_C);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_C);
            } else if (data.equals("d")) {
                r.keyPress(KeyEvent.VK_D);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_D);
            } else if (data.equals("e")) {
                r.keyPress(KeyEvent.VK_E);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_E);
            } else if (data.equals("f")) {
                r.keyPress(KeyEvent.VK_F);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_F);
            } else if (data.equals("g")) {
                r.keyPress(KeyEvent.VK_G);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_G);
            } else if (data.equals("h")) {
                r.keyPress(KeyEvent.VK_H);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_H);
            } else if (data.equals("i")) {
                r.keyPress(KeyEvent.VK_I);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_I);
            } else if (data.equals("j")) {
                r.keyPress(KeyEvent.VK_J);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_J);
            } else if (data.equals("k")) {
                r.keyPress(KeyEvent.VK_K);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_K);
            } else if (data.equals("l")) {
                r.keyPress(KeyEvent.VK_L);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_L);
            } else if (data.equals("m")) {
                r.keyPress(KeyEvent.VK_M);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_M);
            } else if (data.equals("n")) {
                r.keyPress(KeyEvent.VK_N);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_N);
            } else if (data.equals("o")) {
                r.keyPress(KeyEvent.VK_O);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_O);
            } else if (data.equals("p")) {
                r.keyPress(KeyEvent.VK_P);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_P);
            } else if (data.equals("q")) {
                r.keyPress(KeyEvent.VK_Q);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_Q);
            } else if (data.equals("r")) {
                r.keyPress(KeyEvent.VK_R);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_R);
            } else if (data.equals("s")) {
                r.keyPress(KeyEvent.VK_S);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_S);
            } else if (data.equals("t")) {
                r.keyPress(KeyEvent.VK_T);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_T);
            } else if (data.equals("u")) {
                r.keyPress(KeyEvent.VK_U);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_U);
            } else if (data.equals("v")) {
                r.keyPress(KeyEvent.VK_V);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_V);
            } else if (data.equals("w")) {
                r.keyPress(KeyEvent.VK_W);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_W);
            } else if (data.equals("x")) {
                r.keyPress(KeyEvent.VK_X);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_X);
            } else if (data.equals("y")) {
                r.keyPress(KeyEvent.VK_Y);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_Y);
            } else if (data.equals("z")) {
                r.keyPress(KeyEvent.VK_Z);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_Z);
            }else if (data.equals("\\")) {
                r.keyPress(KeyEvent.VK_BACK_SLASH);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_BACK_SLASH);
            }else if (data.equals(":")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_SEMICOLON);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SEMICOLON);
                r.keyRelease(KeyEvent.VK_SHIFT);
            }else if (data.equals("$")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_4);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_4);
                r.keyRelease(KeyEvent.VK_SHIFT);
            }
            //ini untuk CAPSLOCK
            else if (data.equals("A")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_A);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_A);
            } else if (data.equals("B")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_B);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_B);
            } else if (data.equals("C")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_C);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_C);
            } else if (data.equals("D")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_D);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_D);
            } else if (data.equals("E")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_E);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_E);
            } else if (data.equals("F")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_F);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_F);
            } else if (data.equals("G")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_G);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_G);
            } else if (data.equals("H")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_H);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_H);
            } else if (data.equals("I")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_I);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_I);
            } else if (data.equals("J")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_J);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_J);
            } else if (data.equals("K")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_K);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_K);
            } else if (data.equals("L")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_L);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_L);
            } else if (data.equals("M")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_M);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_M);
            } else if (data.equals("N")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_N);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_N);
            } else if (data.equals("O")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_O);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_O);
            } else if (data.equals("P")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_P);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_P);
            } else if (data.equals("Q")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_Q);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_Q);
            } else if (data.equals("R")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_R);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_R);
            } else if (data.equals("S")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_S);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_S);
            } else if (data.equals("T")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_T);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_T);
            } else if (data.equals("U")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_U);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_U);
            } else if (data.equals("V")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_V);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_V);
            } else if (data.equals("W")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_W);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_W);
            } else if (data.equals("X")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_X);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_X);
            } else if (data.equals("Y")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_Y);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_Y);
            } else if (data.equals("Z")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_Z);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SHIFT);
                r.keyRelease(KeyEvent.VK_Z);
            }
            //ini untuk angka
            else if (data.equals("0")) {
                r.keyPress(KeyEvent.VK_0);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_0);
            } else if (data.equals("1")) {
                r.keyPress(KeyEvent.VK_1);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_1);
            } else if (data.equals("2")) {
                r.keyPress(KeyEvent.VK_2);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_2);
            } else if (data.equals("3")) {
                r.keyPress(KeyEvent.VK_3);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_3);
            } else if (data.equals("4")) {
                r.keyPress(KeyEvent.VK_4);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_4);
            } else if (data.equals("5")) {
                r.keyPress(KeyEvent.VK_5);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_5);
            } else if (data.equals("6")) {
                r.keyPress(KeyEvent.VK_6);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_6);
            } else if (data.equals("7")) {
                r.keyPress(KeyEvent.VK_7);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_7);
            } else if (data.equals("8")) {
                r.keyPress(KeyEvent.VK_8);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_8);
            } else if (data.equals("9")) {
                r.keyPress(KeyEvent.VK_9);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_9);
            } else if (data.equals(".")) {
                r.keyPress(KeyEvent.VK_PERIOD);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_PERIOD);
            } else if (data.equals("-")) {
                r.keyPress(KeyEvent.VK_MINUS);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_MINUS);
            } else if (data.equals(" ")) {
                r.keyPress(KeyEvent.VK_SPACE);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SPACE);
            } else if (data.equals("_")) {
                r.keyPress(KeyEvent.VK_TAB);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_TAB);
            } else {
                System.out.println("cannot type that character");
            }
        } catch (Exception e) {
            System.out.println("failed to perform pressKey");
        }
    }

    public static void shortKey(String value){
        try {
            Robot r = new Robot();

            if(value.equalsIgnoreCase("F12")){
                r.keyPress(KeyEvent.VK_F12);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_F12);
                Thread.sleep(100);
                System.out.println("press F12");
            }else if(value.equalsIgnoreCase("END")){
                r.keyPress(KeyEvent.VK_END);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_END);
                Thread.sleep(100);
                System.out.println("press END");
            }else if(value.equalsIgnoreCase("PAGE UP")){
                r.keyPress(KeyEvent.VK_PAGE_UP);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_PAGE_UP);
                Thread.sleep(100);
                System.out.println("press PAGE UP");
            }else if(value.equalsIgnoreCase("ESC")){
                r.keyPress(KeyEvent.VK_ESCAPE);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_ESCAPE);
                Thread.sleep(100);
                System.out.println("press ESC");
            }else if(value.equalsIgnoreCase("ENTER")){
                r.keyPress(KeyEvent.VK_ENTER);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_ENTER);
                Thread.sleep(100);
                System.out.println("press ENTER");
            }else if(value.equalsIgnoreCase("TAB")){
                r.keyPress(KeyEvent.VK_TAB);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_TAB);
                Thread.sleep(100);
                System.out.println("press TAB");
            }else{
                System.out.println("keypress tidak ditemukan");
            }


        } catch (Exception e) {
            System.out.println("failed to press F12");
        }
    }

    public static String ExtractFormulaSheet(String strFormula) {
        String strSheet[] = strFormula.split("!");

        return strSheet[0];
    }

    public static int ExtractFormulaAddress(String strFormula) {

        String strSheet[] = strFormula.split("!");
        CellReference cellReference = new CellReference(strSheet[1]);

        return cellReference.getCol();
    }

    public static String GetCellValue(CellValue cellValue) {

        int cellType = cellValue.getCellType();
        String strResult = "";
        switch (cellType){
            case 0:
                strResult=  cellValue.getNumberValue() + "";
                if (strResult.equalsIgnoreCase("0.0"))
                    strResult = strResult.replace("0.0","");
                else
                    strResult = strResult.replace(".0","");
                break;

            case 1:
                strResult=  cellValue.getStringValue() + "";
                break;

        }
        return strResult;
    }

    public static String ProcessFormulaCell(Workbook workbookTestcase,Row row,Cell cell,int intRow){
        String strResult = "";
        intRow = (intCounterLooping != -1) ? (intStartDataLooping+intCounterLooping )  : intRow ;

        if (row.getCell(5).getCellType() == Cell.CELL_TYPE_FORMULA) {
            FormulaEvaluator evaluator = workbookTestcase.getCreationHelper().createFormulaEvaluator();

            String strFormula = cell.getCellFormula();
            String strSheetDatatable = ExtractFormulaSheet(strFormula);
            int intIndexDatatable = ExtractFormulaAddress(strFormula);
            String strColumn = CellReference.convertNumToColString(intIndexDatatable);

            cell.setCellFormula(strSheetDatatable + "!" + strColumn + intRow);
            CellValue cellValue = evaluator.evaluate(cell);
            strResult = GetCellValue(cellValue);
        } else{
            strResult = cell.toString();
        }

        return strResult;
    }

    public void appendStatusExecuteTable(String[] paramUpdateStatus) throws IOException, InterruptedException {

        String strSheetExecuteTable = "testcase";
        String dataTableFile = paramUpdateStatus[0].replace(" ","%20").replace("\\","/");
        String statusTransaksi = paramUpdateStatus[1];
        String keterangan = paramUpdateStatus[2];
        String strDatatableExecute = paramUpdateStatus[4]+ "/ExecuteTable.xlsx";
        File file = new File(strDatatableExecute);

        //Get the workbook instance for XLSX file
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet spreadsheet = workbook.getSheet(strSheetExecuteTable);
        int intColumnStatus = GetCounterColumn(spreadsheet,"STATUS");
        int intColumnKeterangan =  GetCounterColumn(spreadsheet,"KETERANGAN");
        int intDataKe = Integer.parseInt(paramUpdateStatus[5]);

//        EDIT BY KEH
        CellStyle hyperlinkStyle = workbook.createCellStyle();
        Font hyperlinkFont = workbook.createFont();
        Row row = spreadsheet.getRow(intDataKe);
        Cell cell = row.createCell(intColumnStatus);
        Hyperlink href =  workbook.getCreationHelper().createHyperlink(HyperlinkType.FILE);
        String strAddress = dataTableFile;

        hyperlinkFont.setUnderline(Font.U_SINGLE);
        hyperlinkFont.setColor((statusTransaksi.equalsIgnoreCase("PASSED")) ? IndexedColors.BLUE.getIndex() : IndexedColors.RED.getIndex());
        hyperlinkStyle.setFont(hyperlinkFont);

        href.setAddress(strAddress);
        cell.setHyperlink(href);
        cell.setCellValue(statusTransaksi);
        cell.setCellStyle(hyperlinkStyle);
//        EDIT BY KEH

        spreadsheet.getRow(intDataKe).createCell(intColumnKeterangan).setCellValue(keterangan);

        FileOutputStream outputStream = new FileOutputStream(new File(strDatatableExecute));
        workbook.write(outputStream);
        outputStream.close();


    }

    public static int GetCounterColumn(Sheet sheetTestcase,String strKeyword){
        int i = 0;
        int intResult = -1;
        while(!sheetTestcase.getRow(0).getCell(i).getStringCellValue().equalsIgnoreCase("")){
            String strValue = sheetTestcase.getRow(0).getCell(i).getStringCellValue();
            if(strValue.equalsIgnoreCase(strKeyword)){
                intResult = i;
                break;
            }
            i++;
        }
        return intResult;

    }
}
