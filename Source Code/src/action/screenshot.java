package action;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

import java.io.File;
import java.io.IOException;

public class screenshot {
    public static void take(WebDriver webDriver, File filePath) throws IOException {
        TakesScreenshot takesScreenshot = ((TakesScreenshot) webDriver);
        File sourceFile = takesScreenshot.getScreenshotAs(OutputType.FILE);
        FileUtils.copyFile(sourceFile, filePath);
    }
}
