package log;

import java.io.File;
import java.io.IOException;

public class createLog {
    public static void create(String logFile) throws IOException {
        File file = new File(logFile);
        if (!file.exists()) {
            file.createNewFile();
        }
    }
}
