package report;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class createReport {
    public static void create(String reportFile) throws IOException {
        XWPFDocument document = new XWPFDocument();
        FileOutputStream outputStream = new FileOutputStream(new File(reportFile));
        document.write(outputStream);
        outputStream.close();
    }
}
