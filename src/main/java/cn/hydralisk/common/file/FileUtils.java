package cn.hydralisk.common.file;

import cn.hydralisk.common.constants.CharConstants;

import java.io.*;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.List;

/**
 * 处理文件的工具类
 * @author master.yang
 */
public abstract class FileUtils {

    public static void write(String filePath, String fileName, String content) throws IOException {
        String fileUrl = filePath + fileName;
        File file = new File(filePath);
        if (!file.exists()) {
            file.mkdir();
        }
        FileWriter payFilePW = new FileWriter(fileUrl, false);
        try {
            payFilePW.write(content);
        } finally {
            payFilePW.close();
        }
    }

    public static void readForWrite(String filePath, String fileName, OutputStream outputStream)
                                                                                                throws IOException {
        String fileUrl = filePath + fileName;
        String line = null;
        FileReader fileReader = new FileReader(fileUrl);
        BufferedReader br = new BufferedReader(fileReader);
        line = br.readLine();
        while (line != null) {
            outputStream.write(line.getBytes(Charset.forName("UTF-8")));
            outputStream.write("\r\n".getBytes(Charset.forName("UTF-8")));
            line = br.readLine();
        }
    }

    public static byte[] read(String filePath, String fileName) throws IOException {
        String fileUrl = filePath + fileName;
        String line = null;
        StringBuilder stringBuilder = new StringBuilder();
        FileReader fileReader = new FileReader(fileUrl);
        BufferedReader br = new BufferedReader(fileReader);
        line = br.readLine();
        while (line != null) {
            stringBuilder.append(line);
            line = br.readLine();
        }
        return stringBuilder.toString().getBytes(Charset.forName("UTF-8"));
    }

    public static List<String> readToList(String fileUrl) throws IOException, FileNotFoundException {
        List<String> contents = new ArrayList<String>();
        FileReader fileReader = new FileReader(fileUrl);
        BufferedReader bufferedReader = new BufferedReader((fileReader));
        String lineContent = null;
        while ((lineContent = bufferedReader.readLine()) != null) {
            contents.add(lineContent);
        }
        fileReader.close();
        bufferedReader.close();
        return contents;
    }

    public static List<String> readToList(String fileUrl, String charset) throws IOException,
                                                                         FileNotFoundException {
        List<String> contents = new ArrayList<String>();
        BufferedReader bufferedReader = new BufferedReader(new InputStreamReader(
            new FileInputStream(fileUrl), charset));

        String lineContent = null;
        while ((lineContent = bufferedReader.readLine()) != null) {
            contents.add(lineContent);
        }
        bufferedReader.close();
        return contents;
    }

    public static boolean createFile(String filePath) throws IOException {
        File file = new File(filePath);
        File fileParent = file.getParentFile();
        if (fileParent != null && !fileParent.exists()) {
            boolean mkdirSuccess = fileParent.mkdirs();
            if (!mkdirSuccess) {
                return false;
            }
        }
        if (!file.exists()) {
            boolean createFileSuccess = file.createNewFile();
            if (!createFileSuccess) {
                return false;
            }
        }
        return true;
    }

    /**
     * 获取文件扩展名
     * 例如test.txt结果应该返回txt
     * 
     * @param file
     * @return
     */
    public static String getFileNameExtension(File file) {
        String fileName = file.getName();
        return fileName.substring(fileName.lastIndexOf(CharConstants.POINT) + 1);
    }
}
