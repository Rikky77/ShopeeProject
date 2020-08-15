package report;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class appendReport {
    public static void appendText(String reportFile, String textToAppend) throws IOException {
        File file = new File(reportFile);
        if(!file.exists()) {
            createReport.create(reportFile);
        }

        XWPFDocument document = new XWPFDocument(new FileInputStream(reportFile));
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();

        run.setText(textToAppend);
        document.write(new FileOutputStream(new File(reportFile)));
        //document.close();
    }

    public static void appendImage(String reportFile, File imageFile, String imageText) throws IOException, InvalidFormatException {
        File file = new File(reportFile);
        if(!file.exists()) {
            createReport.create(reportFile);
        }

        XWPFDocument document = new XWPFDocument(new FileInputStream(reportFile));
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(imageText);

        FileInputStream inputStream = new FileInputStream(imageFile.getAbsolutePath());
        run.addPicture(inputStream, XWPFDocument.PICTURE_TYPE_JPEG, imageFile.getAbsolutePath(), Units.toEMU(400), Units.toEMU(200));
        inputStream.close();

        document.write(new FileOutputStream(new File(reportFile)));
        //document.close();
    }
}
