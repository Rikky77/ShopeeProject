package datatable;

import action.MyConfig;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class readDatatable {
    public Sheet readTestcase(String datatablePath, String datatableFile) throws IOException {
        File file = new File(datatablePath + datatableFile);
        FileInputStream inputStream = new FileInputStream(file);
        Workbook workbook;
        String fileExtension = datatableFile.substring(datatableFile.indexOf("."));

        if(fileExtension.equals(".xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        }
        else if(fileExtension.equals(".xls")) {
            workbook = new HSSFWorkbook(inputStream);
        }
        else {
            workbook = null;
        }


        Sheet sheet = workbook.getSheet("testcase");
        return sheet;
    }

    public Sheet readSpecificTestcase(String datatablePathFile, String strSheet) throws IOException {
        File file = new File(datatablePathFile);
        FileInputStream inputStream = new FileInputStream(file);
        Workbook workbook;
        String fileExtension = FilenameUtils.getExtension(datatablePathFile);

        if(fileExtension.equals("xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        }
        else if(fileExtension.equals("xls")) {
            workbook = new HSSFWorkbook(inputStream);
        }
        else {
            workbook = null;
        }

        Sheet sheet = workbook.getSheet(strSheet);
        return sheet;
    }

    public Sheet readDatatable(String datatablePath, String datatableFile) throws IOException {
        File file = new File(datatablePath + datatableFile);
        FileInputStream inputStream = new FileInputStream(file);
        Workbook workbook;
        String fileExtension = datatableFile.substring(datatableFile.indexOf("."));

        if(fileExtension.equals(".xlsx")) {
            workbook = new XSSFWorkbook(inputStream);
        }
        else if(fileExtension.equals(".xls")) {
            workbook = new HSSFWorkbook(inputStream);
        }
        else {
            workbook = null;
        }

        Sheet sheet = workbook.getSheet(MyConfig.strDatatableSheetName);
        return sheet;
    }
}
