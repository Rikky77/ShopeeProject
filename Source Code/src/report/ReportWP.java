package report;

import org.apache.poi.ss.extractor.ExcelExtractor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import testcase.executeTest;

import java.io.*;
import java.util.Arrays;
import java.util.HashMap;

public class ReportWP {

    public ReportWP() {
    }

    public void CreateWPExcel(String strPath, String strTitle, String strType) {


        int intWidth = 15;
        int intHeight = 24;
        int intSpace = 2;
        int intRowDescription = 2;
        int intMasterRow = intRowDescription + intSpace + 1;
        int intMasterColumn = intSpace - 1;

//        Get All File with specific extention
        File f = new File(strPath);

            File[] matchingFiles = f.listFiles(new FilenameFilter() {
            public boolean accept(File dir, String name) {
                return name.endsWith("png");
//                return name.startsWith("xxx") && name.endsWith("jpeg");
            }
        });
        Arrays.sort(matchingFiles);

        /*Create WP*/
        try {

            Workbook wb = new XSSFWorkbook(new FileInputStream(strPath + "WP-" + strTitle + ".xlsx"));
            Sheet sheet = null;

            /*STYLE*/
            XSSFFont headerFont = (XSSFFont) wb.createFont();
            headerFont.setBold(true);
            headerFont.setUnderline(XSSFFont.U_SINGLE);
            headerFont.setFontHeightInPoints((short) 24);

            XSSFFont descriptionFont = (XSSFFont) wb.createFont();
            descriptionFont.setBold(true);
            descriptionFont.setColor(IndexedColors.WHITE.getIndex());

            XSSFCellStyle styleHeader = (XSSFCellStyle) wb.createCellStyle();
            styleHeader.setFont(headerFont);

            XSSFCellStyle styleDescription = (XSSFCellStyle) wb.createCellStyle();
            styleDescription.setFont(descriptionFont);
            styleDescription.setFillForegroundColor(IndexedColors.RED.getIndex());
            styleDescription.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            /*STYLE*/
            int intCodeDataDefault = 0;
            int intCodeDefault = 0;

            for (int i = 0; i < matchingFiles.length; i++) {
                int intCodeData = Integer.parseInt((matchingFiles[i].getName().split("_")[1]).replaceAll("^0+",""));
                int intCode = 0;

                try{
                    intCode = Integer.parseInt((matchingFiles[i].getName().split("_")[2]).replaceAll("^0+",""));
               }catch (Exception ex){

               }

                switch (strType){
                    case "HORIZONTAL":
                        sheet = wb.getSheet("WP");

                        if (sheet == null){
                            sheet = wb.createSheet("WP");
                            Cell celHeader = sheet.createRow(0).createCell(0);
                            celHeader.setCellValue(strTitle);
                            celHeader.setCellStyle(styleHeader);
                        }

                        if (intCodeData > intCodeDataDefault){
                            intCodeDataDefault = intCodeData;
                            intMasterRow = intRowDescription + intSpace + 1;
                            intMasterColumn = intSpace - 1;

                            Row rowDescription = sheet.createRow(intRowDescription);
                            Cell cellDescription = rowDescription.createCell(0);
                            for (int j = 0; j <= 500; j++)
                                rowDescription.createCell(j).setCellStyle(styleDescription);
                            cellDescription.setCellValue("Data ke-" + intCodeData);
                            intRowDescription += (intSpace * 2) + 1 + intHeight;
                        }
                        break;
                    case "VERTICAL":
                        if (intCodeData > intCodeDataDefault){
                            intCodeDataDefault = intCodeData;

                            intMasterRow = intRowDescription + intSpace + 1;
                            intMasterColumn = intSpace - 1;

                            sheet = wb.createSheet("WP-"+intCodeData);
                            Cell celHeader = sheet.createRow(0).createCell(0);
                            celHeader.setCellValue(strTitle);
                            celHeader.setCellStyle(styleHeader);

                            Row rowDescription = sheet.createRow(intRowDescription);
                            Cell cellDescription = rowDescription.createCell(0);

                            for (int j = 0; j <= 500; j++)
                                rowDescription.createCell(j).setCellStyle(styleDescription);
                            cellDescription.setCellValue("Data ke-" + intCodeData);
                        }
                        break;

                    case "HORIZONTAL_DYNAMIC":
                        if (intCodeData > intCodeDataDefault) {
                            intCodeDataDefault = intCodeData;
                            intCodeDefault = 0;
                            sheet = wb.createSheet("WP-" + intCodeData);
                            Cell celHeader = sheet.createRow(0).createCell(0);
                            celHeader.setCellValue(strTitle);
                            celHeader.setCellStyle(styleHeader);

                            intRowDescription = 2;
                            intMasterRow = intRowDescription + intSpace + 1;
                            intMasterColumn = intSpace - 1;

                        }
                        if (intCode > intCodeDefault){
                            intCodeDefault = intCode;
                            intMasterRow = intRowDescription + intSpace + 1;
                            intMasterColumn = intSpace - 1;

                            Row rowDescription = sheet.createRow(intRowDescription);
                            Cell cellDescription = rowDescription.createCell(0);

                            for (int j = 0; j <= 500; j++)
                                rowDescription.createCell(j).setCellStyle(styleDescription);

                            String strDescription = "";
                            try {
                                strDescription = ((HashMap)executeTest.hasCounterDataListDescription.get(intCodeData)).get(intCode).toString();
                            }catch (Exception ex){}
                            cellDescription.setCellValue(strDescription);
                            intRowDescription += (intSpace * 2) + 1 + intHeight;
                        }

                        break;
                    default:break;
                }



                InputStream inputStream = new FileInputStream(matchingFiles[i]);
                byte[] bytes = IOUtils.toByteArray(inputStream);
                int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
                inputStream.close();

                CreationHelper helper = wb.getCreationHelper();
                Drawing drawing = sheet.createDrawingPatriarch();
                ClientAnchor anchor = helper.createClientAnchor();

                //create an anchor with upper left cell _and_ bottom right cell
                anchor.setCol1(intMasterColumn); //Column B
                anchor.setRow1(intMasterRow); //Row 3
                anchor.setCol2(intMasterColumn + intWidth); //Column C
                anchor.setRow2(intMasterRow + intHeight); //Row 4
                if (strType.equalsIgnoreCase("HORIZONTAL") || strType.equalsIgnoreCase("HORIZONTAL_DYNAMIC")) {
                    intMasterColumn += intWidth + intSpace;
                } else if (strType.equalsIgnoreCase("VERTICAL")) {
                    intMasterRow += intHeight + intSpace;
                }

                //Creates a picture
                drawing.createPicture(anchor, pictureIdx);
            }


            //Write the Excel file
            FileOutputStream fileOut = null;
            fileOut = new FileOutputStream(strPath + "WP-" + strTitle + ".xlsx");
            wb.write(fileOut);
            fileOut.close();

        } catch (IOException ioex) {
            System.out.println(ioex.toString());
        }

    }


}
