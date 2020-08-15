package log;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.Timestamp;

public class appendLog {
    public static void append(String logFile, String textToAppend) throws IOException {
        File file = new File(logFile);
        if (!file.exists()) {
            createLog.create(logFile);
        }

        Timestamp timestamp = new Timestamp(System.currentTimeMillis());
        FileWriter fileWriter = new FileWriter(logFile, true);
        BufferedWriter bufferedWriter = new BufferedWriter(fileWriter);

        bufferedWriter.write(timestamp + "   " + textToAppend);
        bufferedWriter.newLine();
        bufferedWriter.close();
    }
}
