package action;

import com.assertthat.selenium_shutterbug.core.Shutterbug;
import com.assertthat.selenium_shutterbug.utils.web.ScrollStrategy;
import datatable.readDatatable;
import log.appendLog;
import objectRepository.readObjectRepository;
import org.apache.commons.io.FileUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.Point;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.w3c.dom.html.HTMLTableCaptionElement;
import org.w3c.dom.html.HTMLTableCellElement;
import report.appendReport;
import result.appendResult;
import testcase.executeTest;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;
import java.io.*;
import java.sql.Timestamp;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;

public class keyword {
    WebDriver webDriver;

    //Variable Capture
    int intCounterData;
    String strAppName = executeTest.strNameApp;
    String strFileName,strDataTableFilePath = "";
    readDatatable objReadDatatable = new readDatatable();

    public keyword(WebDriver webDriver) {
        this.webDriver = webDriver;
    }

    public void perform(String[] param, String logFile, String reportFile, int paramCounterData) throws Exception {
        /*
        param0: keyword
        param1: pagename
        param2: objectname
        param3: objecttype
        param4: value
        */

        this.intCounterData = paramCounterData;
        this.strFileName = param[6];
        this.strDataTableFilePath = param[5];

        //Switch page
        String MainWindow = webDriver.getWindowHandle();
        Boolean flgPage = false;

        Set<String> s1 = webDriver.getWindowHandles();
        Iterator<String> i1 = s1.iterator();

        while (i1.hasNext()) {
            String ChildWindow = i1.next();
            if (!MainWindow.equalsIgnoreCase(ChildWindow)) {
                webDriver.switchTo().window(ChildWindow);
                flgPage = true;
            }
        }


        String strKeyword, strPageName, strObjectName, strObjectType, strValue;
        strKeyword = param[0];
        strPageName = param[1];
        strObjectName = param[2];
        strObjectType = param[3];
        strValue = param[4];


        readObjectRepository objRepository = new readObjectRepository();
        Document xmlObjectRepository = objRepository.read(strPageName);
        try {
            switch (strKeyword.toUpperCase()) {

                case "GOTOURL":
                    webDriver.get(strValue);
                    appendLog.append(logFile, "Go to URL -> " + strValue);
                    appendReport.appendText(reportFile, "Go to URL: " + strValue);
                    break;

                case "SCREENSHOT":
                    ScreenShootExcel(logFile, reportFile);
                    break;

                case "KEYPRESS":
                    Thread.sleep(1000);
                    shortKey(strValue.replace("'",""));
                    appendLog.append(logFile, "KEYPRESS -> "  + ": " + strValue);
                    appendReport.appendText(reportFile, "KEYPRESS : " + " - " + strValue);
                    break;

                case "SETTEXT":
                    loopWhile(xmlObjectRepository, strObjectName, strObjectType);
                    if (!strValue.trim().equals("")) {
                        webDriver.findElement(getObjectBy(xmlObjectRepository, strObjectName, strObjectType)).clear();
                    }

                    if(strValue.contains("[")) {
                        strValue = executeTest.mapValueTemp.get(strValue);
                    }
                    webDriver.findElement(getObjectBy(xmlObjectRepository, strObjectName, strObjectType)).sendKeys(strValue);
                    appendLog.append(logFile, "Set text -> " + strObjectName + ": " + strValue);
                    appendReport.appendText(reportFile, "Set text: " + strObjectName + " - " + strValue);
                    break;

                case "CLICK":
                    webDriver.manage().timeouts().implicitlyWait(1, TimeUnit.SECONDS);
                    loopWhile(xmlObjectRepository, strObjectName, strObjectType);

                    webDriver.findElement(getObjectBy(xmlObjectRepository, strObjectName, strObjectType)).click();
                    appendLog.append(logFile, "Click -> " + strObjectName + ": " + strValue);
                    appendReport.appendText(reportFile, "Click: " + strObjectName);
                    break;

                case "CLICK_TEXT":
                    webDriver.manage().timeouts().implicitlyWait(1, TimeUnit.SECONDS);
                    Thread.sleep(3000);

                    webDriver.findElement(By.xpath("//*[text()=\""+strValue+"\"]")).click();
                    appendLog.append(logFile, "CLICK_TEXT -> " + ": " + strValue);
                    appendReport.appendText(reportFile, "CLICK_TEXT: " + strValue);
                    break;

                case "SELECT":
                    loopWhile(xmlObjectRepository, strObjectName, strObjectType);
                    String strList;
                    Select dropdown = new Select(webDriver.findElement(getObjectBy(xmlObjectRepository, strObjectName, strObjectType)));
                    //Get all options
                    List<WebElement> lstDropdown = dropdown.getOptions();
                    //Get the length
                    // Loop to print one by one
                    for (int j = 0; j < lstDropdown.size(); j++) {
                        strList = lstDropdown.get(j).getText();
                        if (strList.trim().equals(strValue.trim())) {
                            dropdown.selectByIndex(j);
                            break;
                        }
                    }
                    appendLog.append(logFile, "Select -> " + strObjectName + ": " + strValue);
                    appendReport.appendText(reportFile, "Select: " + strObjectName);
                    break;
            }
            if (flgPage = true) {
                webDriver.switchTo().window(MainWindow);
            }
            String[] paramUpdateStatus = new String[4];
            paramUpdateStatus[0] = strDataTableFilePath;
            paramUpdateStatus[1] = "PASSED";
            paramUpdateStatus[2] = "";
            paramUpdateStatus[3] = String.valueOf(intCounterData);
            appendStatusDT(paramUpdateStatus);

        } catch (Exception e) {
            System.out.println(e.toString());
            executeTest.bolValidationNextStep = false;

            ScreenShootExcel(logFile, reportFile);

            String[] paramUpdateStatus = new String[4];
            paramUpdateStatus[0] = strDataTableFilePath;
            paramUpdateStatus[1] = "FAILED";
            paramUpdateStatus[2] = "ERROR Perform -> " + strKeyword + ", "+ strObjectName + ", " +
                    strValue +" Error Line : "+ (executeTest.intCounterRow +1) + "\n" + e.toString();
            paramUpdateStatus[3] = String.valueOf(intCounterData);
            appendStatusDT(paramUpdateStatus);

            //logout

            appendLog.append(logFile, "ERROR Perform -> " + strKeyword + ", "+ strObjectName + ", " + strValue);
            appendReport.appendText(reportFile, "ERROR Perform " + strKeyword + ", "+ strObjectName + ", " + strValue);

        }
    }

    public void expected(String[] param, Row rowExpected ) throws Exception {
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

        String strKeyword, strPageName, strObjectName, strObjectType, strValue, logFile, reportFile, resultFile, testCase, strDataTableFilePath, strDateKeReport;
        strKeyword = param[0];
        strPageName = param[1];
        strObjectName = param[2];
        strObjectType = param[3];
        strValue = param[4];
        logFile = param[5];
        reportFile = param[6];
        resultFile = param[7];
        testCase = param[8];
        strDataTableFilePath = param[9];
        strDateKeReport = param[10];

        String strActualValue, strStatus;
        String[] paramExpected = new String[5];
        paramExpected[1] = testCase;
        paramExpected[2] = strValue;
        paramExpected[4] = strDateKeReport;

        readObjectRepository objRepository = new readObjectRepository();
        Document xmlObjectRepository = objRepository.read(strPageName);
        try {
            switch (strKeyword.toUpperCase()) {
                case "COMPARENULL" :
                    Integer intTotalRow = getTotalElement(xmlObjectRepository, strObjectName, strObjectType);

                    String strTglEff, strTglPost, strStatusTgl;
                    String xptTableRowTglEfektif, xptTableRowTglPost;
                    strStatusTgl = "";
                    for (int z = 1; z <= intTotalRow; z++) {
                        xptTableRowTglPost ="//*[@id=\"adjust-extraction\"]/tbody/tr["+ z +"]/td[2]/div/div[2]/div/div/input";
                        xptTableRowTglEfektif = "//*[@id=\"adjust-extraction\"]/tbody/tr["+ z +"]/td[3]/div/div[2]/div/div/input";
                        strTglEff = webDriver.findElement(By.xpath(xptTableRowTglEfektif)).getText();
                        strTglPost = webDriver.findElement(By.xpath(xptTableRowTglPost)).getText();

                        if ((strTglPost.trim().toUpperCase()).equals("NULL")) {
                            strStatusTgl = "TglPostFailed-row-" +z;
                        }

                        if ((strTglEff.trim().toUpperCase()).equals("NULL")) {
                            strStatusTgl = strStatusTgl + " & TglEffFailed-row" +z;
                        }

                    }

                    if(strStatusTgl.equals("")){
                        strStatusTgl = "TglPostPassed & TglEffPassed";
                    }

                    strActualValue = strStatusTgl;

                    if(strStatusTgl.equals("TglPostPassed & TglEffPassed")){
                        strStatus = "PASSED";
                    }else{
                        strStatus = "FAILED";
                    }

                        /*
                        param0: Status
                        param1: TestCase
                        param2: Expected
                        param3: Actual
                        */
                    paramExpected[0] = strStatus;
                    paramExpected[3] = strActualValue;

                    if(strStatus.equals("FAILED")){
                        String[] paramUpdateStatus = new String[4];

                        paramUpdateStatus[0] = strDataTableFilePath;
                        paramUpdateStatus[1] = strStatus;
                        paramUpdateStatus[2] = "Error 3 : Total debit & credit sesuai tanggal ada yang NULL";
                        paramUpdateStatus[3] = String.valueOf(intCounterData);
                        appendStatusDT(paramUpdateStatus);
                    }else{
                        String[] paramUpdateStatus = new String[4];

                        paramUpdateStatus[0] = strDataTableFilePath;
                        paramUpdateStatus[1] = strStatus;
                        paramUpdateStatus[2] = "";
                        paramUpdateStatus[3] = String.valueOf(intCounterData);
                        appendStatusDT(paramUpdateStatus);
                    }

                    appendResult.appendCell(paramExpected, resultFile);
                    appendLog.append(logFile, "Expected - COMPARENULL -> " + strActualValue + ": " + strObjectName +  ":" + strStatus);
                    appendReport.appendText(reportFile, "Status: " + strStatus);
                    break;
                case "COMPAREDEBITKREDIT":
                    String debitTxn, debitFile, creditTxn, creditFile, strStatusTemp;

                    debitTxn = executeTest.mapValueTemp.get("strDebitTxn");
                    debitFile =  executeTest.mapValueTemp.get("strDebitTxn");
                    creditTxn = executeTest.mapValueTemp.get("strCreditTxn");
                    creditFile = executeTest.mapValueTemp.get("strCreditFile");

                    if(debitTxn.equals(debitFile)){
                        strStatusTemp = "DebitPassed";
                    }else {
                        strStatusTemp = "DebitFailed";
                    }

                    if(creditTxn.equals(creditFile)){
                        strStatusTemp = strStatusTemp + " & CreditPassed";
                    }else {
                        strStatusTemp = strStatusTemp + " & CreditPassed";
                    }

                    if (strStatusTemp.equals("DebitPassed & CreditPassed")) {
                        strStatus = "PASSED";
                    }else {
                        strStatus = "FAILED";
                    }

                    strActualValue = strStatusTemp;


                        /*
                        param0: Status
                        param1: TestCase
                        param2: Expected
                        param3: Actual
                        */
                    paramExpected[0] = strStatus;
                    paramExpected[3] = strActualValue;

                    if(strStatus.equals("FAILED")){
                        String[] paramUpdateStatus = new String[4];

                        paramUpdateStatus[0] = strDataTableFilePath;
                        paramUpdateStatus[1] = strStatus;
                        paramUpdateStatus[2] = "Error 4 : Total debit & credit tidak sesuai";
                        paramUpdateStatus[3] = String.valueOf(intCounterData);
                        appendStatusDT(paramUpdateStatus);
                    }else{
                        String[] paramUpdateStatus = new String[4];

                        paramUpdateStatus[0] = strDataTableFilePath;
                        paramUpdateStatus[1] = strStatus;
                        paramUpdateStatus[2] = "";
                        paramUpdateStatus[3] = String.valueOf(intCounterData);
                        appendStatusDT(paramUpdateStatus);
                    }

                    appendResult.appendCell(paramExpected, resultFile);
                    appendLog.append(logFile, "Expected - COMPAREDEBITKREDIT -> " + strActualValue + ": " + strObjectName +  ":" + strStatus);
                    appendReport.appendText(reportFile, "Status: " + strStatus);
                    break;
                case "EXIST":
                    loopWhile(xmlObjectRepository, strObjectName, strObjectType);
                    boolean bolElementIsExist = webDriver.findElement(getObjectBy(xmlObjectRepository, strObjectName, strObjectType)).isDisplayed();
                    strActualValue = String.valueOf(bolElementIsExist).trim().toUpperCase();
                    strStatus = GetStatus(strValue, strActualValue);

                        /*
                        param0: Status
                        param1: TestCase
                        param2: Expected
                        param3: Actual
                        */
                    paramExpected[0] = strStatus;
                    paramExpected[3] = strActualValue;
                    if (strObjectName.equals("btnSaveAsDraft")) {
                        String[] paramUpdateStatus = new String[4];

                        paramUpdateStatus[0] = strDataTableFilePath;
                        paramUpdateStatus[1] = strStatus;

                        paramUpdateStatus[3] = String.valueOf(intCounterData);
                        if(strStatus.equals("FAILED")){
                            paramUpdateStatus[2] = "Error 2 : Done tapi tidak bisa diview";
                        }else {
                            paramUpdateStatus[2] = "";
                        }
                        appendStatusDT(paramUpdateStatus);
                    }

                    appendResult.appendCell(paramExpected, resultFile);
                    appendLog.append(logFile, "Expected - EXIST -> " + strActualValue + ": " + strObjectName);
                    appendReport.appendText(reportFile, "Status: " + strStatus);
                    break;

                case "GETTEXT":
                    loopWhile(xmlObjectRepository, strObjectName, strObjectType);

                    strActualValue = (webDriver.findElement(getObjectBy(xmlObjectRepository, strObjectName, strObjectType)).getText().trim().toUpperCase());
                    strStatus = GetStatus(strValue, strActualValue);
                       /*
                        param0: Status
                        param1: TestCase
                        param2: Expected
                        param3: Actual
                        */
                    paramExpected[0] = strStatus;
                    paramExpected[3] = strActualValue;
                    if(strObjectName.equals("lblStatus")) {
                        String[] paramUpdateStatus = new String[4];

                        paramUpdateStatus[0] = strDataTableFilePath;
                        paramUpdateStatus[1] = strStatus;

                        paramUpdateStatus[3] = String.valueOf(intCounterData);
                        if (strStatus.equals("FAILED")) {
                            paramUpdateStatus[2] = "Error 1 : Failed";

                        }else{
                            paramUpdateStatus[2] = "";
                        }

                        appendStatusDT(paramUpdateStatus);
                    }

                    appendResult.appendCell(paramExpected, resultFile);
                    appendLog.append(logFile, "Expected - GETTEXT -> " + strActualValue + ": " + strObjectName);
                    appendReport.appendText(reportFile, "Status: " + strStatus);
                    break;

                case "ISENABLE":
                    loopWhile(xmlObjectRepository, strObjectName, strObjectType);
                    boolean bolElementIsEnable = webDriver.findElement(getObjectBy(xmlObjectRepository, strObjectName, strObjectType)).isEnabled();
                    strActualValue = String.valueOf(bolElementIsEnable).trim().toUpperCase();
                    strStatus = GetStatus(strValue, strActualValue);

                        /*
                        param0: Status
                        param1: TestCase
                        param2: Expected
                        param3: Actual
                        */
                    paramExpected[0] = strStatus;
                    paramExpected[3] = strActualValue;

                    appendResult.appendCell(paramExpected, resultFile);
                    appendLog.append(logFile, "Expected - ISENABLE  -> " + strActualValue + ": " + strObjectName);
                    appendReport.appendText(reportFile, "Status: " + strStatus);
                    break;

                case "GETATRIBUTE":
                    loopWhile(xmlObjectRepository, strObjectName, strObjectType);

                    strActualValue = (webDriver.findElement(getObjectBy(xmlObjectRepository, strObjectName, strObjectType)).getAttribute("value").trim().toUpperCase());

                    strStatus = GetStatus(strValue, strActualValue);
                       /*
                        param0: Status
                        param1: TestCase
                        param2: Expected
                        param3: Actual
                        */
                    paramExpected[0] = strStatus;
                    paramExpected[3] = strActualValue;

                    appendResult.appendCell(paramExpected, resultFile);
                    appendLog.append(logFile, "Expected - GETTEXT -> " + strActualValue + ": " + strObjectName);
                    appendReport.appendText(reportFile, "Status: " + strStatus);
                    break;
            }

        } catch (Exception e) {
            System.out.println(e);
            strActualValue = "Object Not Found";
            strStatus = GetStatus(strValue, strActualValue);
                       /*
                        param0: Status
                        param1: TestCase
                        param2: Expected
                        param3: Actual
                        */
            paramExpected[0] = strStatus;
            paramExpected[1] = testCase;
            paramExpected[2] = strValue;
            paramExpected[3] = strActualValue;

            appendResult.appendCell(paramExpected, resultFile);
            appendLog.append(logFile, "ERROR Expected - " + strObjectName + " " + strKeyword + " - Object Not Fo//'und");
            appendReport.appendText(reportFile, "ERROR Expected - " + strObjectName + " " + strKeyword + " - Object Not Found");
        }
    }


    private By getObjectBy(Document xmlObjectRepository, String objectName, String objectType) throws Exception {
        if (objectType.equalsIgnoreCase("XPATH")) {
            return By.xpath(getObjectXml(xmlObjectRepository, objectName));
        } else {
            throw new Exception("Wrong object type");
        }
    }

    private void ScreenShootExcel(String logFile, String reportFile) throws Exception {
        Thread.sleep(1000);
        String strImagePath = new File(logFile).getParent() + "\\" + strFileName;
        String strImageFileName = strAppName + "_" +
                ("000" + Integer.toString(intCounterData)).substring(("000" + Integer.toString(intCounterData)).length() - 3) + "_" +
                ("000" + Integer.toString(executeTest.intSpecificWPCounter)).substring(("000" + Integer.toString(executeTest.intSpecificWPCounter)).length() - 3) + "_" +
                ("000" + Integer.toString(executeTest.intWPCounter)).substring(("000" + Integer.toString(executeTest.intWPCounter)).length() - 3);

//        Shutterbug.shootPage(webDriver, ScrollStrategy.BOTH_DIRECTIONS,500).withName(strImageFileName).save(strImagePath);
        Shutterbug.shootPage(webDriver, ScrollStrategy.WHOLE_PAGE,500).withName(strImageFileName).save(strImagePath);
        executeTest.intWPCounter += 1;

        appendLog.append(logFile, "Screenshot -> " + strImageFileName + ".png");


    }

    private void ScreenShootElement(String logFile, String reportFile,Document xmlObjectRepository,String strObjectName, String strObjectType) throws Exception {
        Thread.sleep(1000);

        String strImagePath = new File(logFile).getParent() + "\\" + strFileName;

        JavascriptExecutor jsExec = (JavascriptExecutor) webDriver;
        Long webpageHeight = (Long) jsExec.executeScript("return document.body.scrollHeight;"); //get Value Scroll Height
        Integer intTotalPgsDown = 0;

        WebElement ele = webDriver.findElement(getObjectBy(xmlObjectRepository, strObjectName, strObjectType));
        Long intHeight = (Long) jsExec.executeScript("return arguments[0].scrollHeight;",ele);


        if ((intHeight >= webpageHeight)) {
            intTotalPgsDown = Math.round(intHeight / webpageHeight);
            for (int i = 0; i <= intTotalPgsDown; i++) {

                String strImageFileName = strAppName + "_" +
                        ("000" + Integer.toString(intCounterData)).substring(("000" + Integer.toString(intCounterData)).length() - 3) + "_" +
                        ("000" + Integer.toString(executeTest.intSpecificWPCounter)).substring(("000" + Integer.toString(executeTest.intSpecificWPCounter)).length() - 3) + "_" +
                        ("000" + Integer.toString(executeTest.intWPCounter)).substring(("000" + Integer.toString(executeTest.intWPCounter)).length() - 3);

                //ini frame aja
                Shutterbug.shootElement(webDriver, ele).withName(strImageFileName).save(strImagePath);
                //ini satu pages
                //Shutterbug.shootPage(webDriver, ScrollStrategy.WHOLE_PAGE,500).withName(strImageFileName).save(strImagePath);

                executeTest.intWPCounter += 1;
                keyPageDownElement(intHeight,ele);
            }
        } else {
            ScreenShootExcel(logFile,reportFile);
        }


    }


    private void ScreenShootLCAutomice(String logFile,String reportFile,Document xmlObjectRepository,String strObjectName, String strObjectType) throws Exception {
        if(!MyConfig.strTempScenarioBefore.equalsIgnoreCase(MyConfig.strTempScenario)){
            MyConfig.strTempScenarioBefore = MyConfig.strTempScenario;

            ScreenShootElement(logFile, reportFile,xmlObjectRepository,strObjectName,strObjectType);

        }

    }


    private void ScreenShotNoScrolling(String logFile, String reportFile) throws Exception {
        Thread.sleep(1000);
        File imageFile = new File(new File(logFile).getParent() + "\\" + strFileName + "\\" + strAppName + "_" +
                ("000" + Integer.toString(intCounterData)).substring(("000" + Integer.toString(intCounterData)).length() - 3) + "_" +
                ("000" + Integer.toString(executeTest.intSpecificWPCounter)).substring(("000" + Integer.toString(executeTest.intSpecificWPCounter)).length() - 3) + "_" +
                ("000" + Integer.toString(executeTest.intWPCounter)).substring(("000" + Integer.toString(executeTest.intWPCounter)).length() - 3) + ".jpeg");

        screenshot.take(webDriver, imageFile);
        executeTest.intWPCounter += 1;


        appendLog.append(logFile, "Screenshot -> " + imageFile.getName());
        appendReport.appendImage(reportFile, imageFile, "Screenshot: " + imageFile.getName());


    }


    private void scrollingCapture(String[] paramCapture, Long height) throws Exception {
        Integer TotalPgsDown = 0;
        File imageFileScroll;
        Timestamp timestamp;
        /*
        paramCapture0: objectname
        paramCapture1: objecttype
        paramCapture2: logFile
        paramCapture3: TotalPageDown
        paramCapture4: reportFile
        */


        String objectName, objectType, logFile, reportFile;
        Integer TotalPageDown;
        logFile = paramCapture[0];
        TotalPageDown = Integer.parseInt(paramCapture[1]);
        reportFile = paramCapture[2];
        timestamp = new Timestamp(System.currentTimeMillis());
        imageFileScroll = new File(new File(logFile).getParent() + "\\" + strFileName + "\\" + strAppName + "_" +
                ("000" + Integer.toString(intCounterData)).substring(("000" + Integer.toString(intCounterData)).length() - 3) + "_" +
                ("000" + Integer.toString(executeTest.intSpecificWPCounter)).substring(("000" + Integer.toString(executeTest.intSpecificWPCounter)).length() - 3) + "_" +
                ("000" + Integer.toString(executeTest.intWPCounter)).substring(("000" + Integer.toString(executeTest.intWPCounter)).length() - 3) + ".png");

        screenshot.take(webDriver, imageFileScroll);

        appendLog.append(logFile, "Screenshot -> " + imageFileScroll.getName());
        appendReport.appendImage(reportFile, imageFileScroll, "Screenshot: " + imageFileScroll.getName());
        executeTest.intWPCounter += 1;

        for (int i = 0; i < TotalPageDown; i++) {
            keyPageDown(height);
            timestamp = new Timestamp(System.currentTimeMillis());
            imageFileScroll = new File(new File(logFile).getParent() + "\\" + strFileName + "\\" + strAppName + "_" +
                    ("000" + Integer.toString(intCounterData)).substring(("000" + Integer.toString(intCounterData)).length() - 3) + "_" +
                    ("000" + Integer.toString(executeTest.intSpecificWPCounter)).substring(("000" + Integer.toString(executeTest.intSpecificWPCounter)).length() - 3) + "_" +
                    ("000" + Integer.toString(executeTest.intWPCounter)).substring(("000" + Integer.toString(executeTest.intWPCounter)).length() - 3) + ".png");


            screenshot.take(webDriver, imageFileScroll);

            appendLog.append(logFile, "Screenshot Scroll-> " + imageFileScroll.getName() + " Pages : " + Math.round(i + 2));
            appendReport.appendImage(reportFile, imageFileScroll, "Screenshot Scroll: " + imageFileScroll.getName());
            executeTest.intWPCounter += 1;
        }
        //====================================================== EDIT BY KEH ======================================================

        keyHome(height);
    }

    private void keyPageDown(Long height) throws Exception {
        Thread.sleep(250);
        JavascriptExecutor js = (JavascriptExecutor) webDriver;
        js.executeScript("window.scrollBy(0," + Math.round(height - 10) + ");");
    }

    private void keyHome(Long height) throws Exception {
        Thread.sleep(250);

        JavascriptExecutor js = (JavascriptExecutor) webDriver;
        Long value = (Long) js.executeScript("return window.pageYOffset;");
        if (value != 0) {
            js.executeScript("window.scrollBy(0," + Math.round(0 - value) + ");");
        }
    }

    private void keyPageDownElement(Long height,WebElement webElement) throws Exception {
        Thread.sleep(250);
        JavascriptExecutor js = (JavascriptExecutor) webDriver;
        js.executeScript("arguments[0].scrollBy(0," + Math.round(height - 10) + ");",webElement);
    }

    private void keyHomeElement(Long height,WebElement webElement) throws Exception {
        Thread.sleep(250);

        JavascriptExecutor js = (JavascriptExecutor) webDriver;
        Long value = (Long) js.executeScript("return window.pageYOffset;");
        if (value != 0) {
            js.executeScript("arguments[0].scrollBy(0," + Math.round(0 - value) + ");",webElement);
        }
    }

    public boolean isScroll() throws Exception {
        String execScript = "return document.documentElement.scrollHeight>document.documentElement.clientHeight;";
        JavascriptExecutor scrollBarPresent = (JavascriptExecutor) webDriver;
        Boolean bolScroll = (Boolean) (scrollBarPresent.executeScript(execScript));

        return bolScroll;
    }

    private void loopWhile(Document xmlObjectRepository, String objectName, String objectType) throws Exception {
        // boolean bolObjectExist, bolObjectEnable;
        // bolObjectExist = false;
        boolean bolObjectEnable = false;
        int intObjectExist = 0;

        FluentWait<WebDriver> fluentWait = new FluentWait<>(webDriver)
                .withTimeout(30, TimeUnit.SECONDS)
                .pollingEvery(200, TimeUnit.MILLISECONDS)
                .ignoring(NoSuchElementException.class);

        int intFlag = 0;
        do {
            intObjectExist = webDriver.findElements(getObjectBy(xmlObjectRepository, objectName, objectType)).size();

            if (intObjectExist > 0) {
                bolObjectEnable = webDriver.findElement(getObjectBy(xmlObjectRepository, objectName, objectType)).isEnabled();
            } else {
                bolObjectEnable = false;
            }

            if (intFlag == 5) {
                intObjectExist = 1;
                bolObjectEnable = true;
            }

            Thread.sleep(100);

            intFlag = intFlag + 1;
        } while (bolObjectEnable == false);
    }

    private Integer getTotalElement(Document xmlObjectRepository, String objectName, String objectType) throws Exception {

        Integer intTotalPages;
        intTotalPages = webDriver.findElements(getObjectBy(xmlObjectRepository, objectName, objectType)).size();
        return intTotalPages;
    }

    private String getObjectXml(Document xmlObjectRepository, String objectName) {
        String strTextContent = "";

        NodeList nodeList = xmlObjectRepository.getElementsByTagName(objectName);
        for (int i = 0; i < nodeList.getLength(); i++) {
            Node node = nodeList.item(i);
            if (node.getNodeType() == Node.ELEMENT_NODE) {
                NodeList childList = node.getChildNodes();
                for (int j = 0; j < childList.getLength(); j++) {
                    Node childNode = childList.item(j);
                    strTextContent = childNode.getTextContent();
                }
            }
        }

        return strTextContent;
    }

    private String GetStatus(String valueExpected, String actualValue) {
        String strStatus = "";
        if (actualValue.contains(valueExpected.trim().toUpperCase())) {
            strStatus = "PASSED";
        } else {
            strStatus = "FAILED";
        }

        return strStatus;
    }


    public static void emojiSettextrobot(String data) throws AWTException {
        inputKey(data);
        Robot r = new Robot();
        r.keyPress(KeyEvent.VK_ENTER);
        r.delay(100);
        r.keyRelease(KeyEvent.VK_ENTER);
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
            }else if(value.equalsIgnoreCase("ALTF4")){
                r.keyPress(KeyEvent.VK_ALT);
                r.keyPress(KeyEvent.VK_F4);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_TAB);
                r.keyRelease(KeyEvent.VK_F4);
                Thread.sleep(100);
                System.out.println("press TAB");
            }else if(value.equalsIgnoreCase("HOME")){
                r.keyPress(KeyEvent.VK_HOME);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_HOME);
                Thread.sleep(100);
                System.out.println("press HOME");
            }else if(value.equalsIgnoreCase("PAGE DOWN")){
                r.keyPress(KeyEvent.VK_PAGE_DOWN);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_PAGE_DOWN);
                Thread.sleep(100);
                System.out.println("press PAGE DOWN");
            }else if(value.equalsIgnoreCase("RIGHT")) {
                r.keyPress(KeyEvent.VK_RIGHT);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_RIGHT);
                Thread.sleep(100);
                System.out.println("press RIGHT");
            }else{
                System.out.println("keypress tidak ditemukan");
            }


        } catch (Exception e) {
            System.out.println("failed to press F12");
        }
    }

    public static void inputKey(String data) {
        try {
            String[] arr = data.replace("0.0","").split("");
            for (Integer i = 0; i < arr.length; i++) {
                //panggil keyPress
                keyPress(arr[i]);
            }
        } catch (Exception e) {
            System.out.println("failed to perform inputKey");
        }
    }

    public static void keyPress(String data) throws AWTException {

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
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_MINUS);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_MINUS);
                r.keyRelease(KeyEvent.VK_SHIFT);
            } else if (data.equals(":")) {
                r.keyPress(KeyEvent.VK_SHIFT);
                r.keyPress(KeyEvent.VK_SEMICOLON);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_SEMICOLON);
                r.keyRelease(KeyEvent.VK_SHIFT);
            } else if (data.equals("\\")) {
                r.keyPress(KeyEvent.VK_BACK_SLASH);
                r.delay(100);
                r.keyRelease(KeyEvent.VK_BACK_SLASH);
            } else {
                System.out.println("cannot type that character");
            }

        } catch (Exception e) {
            System.out.println("failed to perform pressKey");
        }
    }

    private void ScreenShootTake(String logFile, String reportFile) throws Exception {
        Thread.sleep(1000);
        JavascriptExecutor jsExec = (JavascriptExecutor) webDriver;
        Long webpageHeight = (Long) jsExec.executeScript("return document.body.scrollHeight;"); //get Value Scroll Height
        Integer intTotalPgsDown = 0;
        Timestamp timestamp = new Timestamp(System.currentTimeMillis());
        Integer intHeight = webDriver.manage().window().getSize().height; //get Value window height
        if ((webpageHeight >= intHeight)) {
            keyHome(intHeight);
            intTotalPgsDown = Math.round(webpageHeight / intHeight);
            String[] paramCapture = new String[3];
            paramCapture[0] = logFile.toString();
            paramCapture[1] = intTotalPgsDown.toString();
            paramCapture[2] = reportFile;

            scrollingCapture(paramCapture, intHeight);
        } else {



            File imageFile = new File(new File(logFile).getParent() +"\\" + strFileName+"\\"+  strAppName + "_"+ ("000"+Integer.toString(intCounterData)).substring(("000"+Integer.toString(intCounterData)).length()-3)  +"_" +  ("000"+Integer.toString(executeTest.intWPCounter)).substring(("000" + Integer.toString(executeTest.intWPCounter)).length()-3) + ".jpeg");
            screenshot.take(webDriver, imageFile);
            executeTest.intWPCounter += 1;


            appendLog.append(logFile, "Screenshot -> " + imageFile.getName());
            appendReport.appendImage(reportFile, imageFile, "Screenshot: " + imageFile.getName());
        }


    }


    private void scrollingCapture(String[] paramCapture, Integer height) throws Exception {
        Integer TotalPgsDown = 0;
        File imageFileScroll;
        Timestamp timestamp;
        /*
        paramCapture0: objectname
        paramCapture1: objecttype
        paramCapture2: logFile
        paramCapture3: TotalPageDown
        paramCapture4: reportFile
        */


        String objectName, objectType, logFile, reportFile;
        Integer TotalPageDown;
        logFile = paramCapture[0];
        TotalPageDown = Integer.parseInt(paramCapture[1]);
        reportFile = paramCapture[2];
        timestamp = new Timestamp(System.currentTimeMillis());
        imageFileScroll = new File(new File(logFile).getParent() +"\\" + strFileName+"\\"+  strAppName + "_"+ ("000"+Integer.toString(intCounterData)).substring(("000"+Integer.toString(intCounterData)).length()-3) +"_" +  ("000"+Integer.toString(executeTest.intWPCounter)).substring(("000" + Integer.toString(executeTest.intWPCounter)).length()-3)  + ".jpeg");
        keyHome(height);
        screenshot.take(webDriver, imageFileScroll);

        appendLog.append(logFile, "Screenshot -> " + imageFileScroll.getName());
        appendReport.appendImage(reportFile, imageFileScroll, "Screenshot: " + imageFileScroll.getName());
        executeTest.intWPCounter += 1;

        for (int i = 0; i < TotalPageDown; i++) {
            keyPageDown(height);
            timestamp = new Timestamp(System.currentTimeMillis());
            imageFileScroll = new File(new File(logFile).getParent() +"\\" + strFileName+"\\"+  strAppName + "_"+ ("000"+Integer.toString(intCounterData)).substring(("000"+Integer.toString(intCounterData)).length()-3) +"_" + ("000"+Integer.toString(executeTest.intWPCounter)).substring(("000" + Integer.toString(executeTest.intWPCounter)).length()-3) + ".jpeg");
            screenshot.take(webDriver, imageFileScroll);

            appendLog.append(logFile, "Screenshot Scroll-> " + imageFileScroll.getName() + " Pages : " + Math.round(i + 2));
            appendReport.appendImage(reportFile, imageFileScroll, "Screenshot Scroll: " + imageFileScroll.getName());
            executeTest.intWPCounter += 1;
        }
        //====================================================== EDIT BY KEH ======================================================

        //keyHome(height);
    }
    private void keyPageDown(Integer height) throws Exception {
        Thread.sleep(250);
        JavascriptExecutor js = (JavascriptExecutor) webDriver;
        js.executeScript("window.scrollBy(0,"+ Math.round(height-10) +");");
    }

    private void keyHome(Integer height) throws Exception {
        Thread.sleep(250);

        JavascriptExecutor js = (JavascriptExecutor) webDriver;
        Long value = (Long) js.executeScript("return window.pageYOffset;");
        if(value != 0){
            js.executeScript("window.scrollBy(0,"+ Math.round(0-value) +");");
        }
    }

    public static void appendStatusDT(String[] paramUpdateStatus) throws IOException, InterruptedException {

        String dataTableFile, statusTransaksi, keterangan;
        int dataKe;
        dataTableFile = paramUpdateStatus[0];
        statusTransaksi = paramUpdateStatus[1];
        keterangan = paramUpdateStatus[2];
        dataKe = Integer.parseInt(paramUpdateStatus[3]);
        File file = new File(dataTableFile);

        //Get the workbook instance for XLSX file
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(file));
        //get spreadsheet
        XSSFSheet spreadsheet = workbook.getSheet(MyConfig.strDatatableSheetName);

        int intDataKe = dataKe;
        int intColumnStatus = executeTest.GetCounterColumn(spreadsheet,"STATUS");
        int intColumnKeterangan = executeTest.GetCounterColumn(spreadsheet,"KETERANGAN");

//        EDIT BY KEH UNTUK HYPERLINK
//        CellStyle hyperlinkStyle = workbook.createCellStyle();
//        Font hyperlinkFont = workbook.createFont();
//        Row row = spreadsheet.getRow(intDataKe);
//        Cell cell = row.createCell(MyConfig.intColumnStatus);
//        Hyperlink href =  workbook.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
//        String strAddress = "=FILTER!A1";
//
//        hyperlinkFont.setUnderline(Font.U_SINGLE);
//        hyperlinkFont.setColor((statusTransaksi.equalsIgnoreCase("PASSED")) ? IndexedColors.BLUE.getIndex() : IndexedColors.RED.getIndex());
//        hyperlinkStyle.setFont(hyperlinkFont);
//
//        href.setAddress(strAddress);
//        cell.setHyperlink(href);
//        cell.setCellValue(statusTransaksi);
//        cell.setCellStyle(hyperlinkStyle);
//        EDIT BY KEH

        spreadsheet.getRow(intDataKe).createCell(intColumnStatus).setCellValue(statusTransaksi);
        spreadsheet.getRow(intDataKe).createCell(intColumnKeterangan).setCellValue(keterangan);

        FileOutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();

    }
}
