package cn.hydralisk.common.file;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 * 解压文件
 * @author master.yang
 */
public class UncompressionFileUtils {

    private static final Logger logger = LoggerFactory.getLogger(UncompressionFileUtils.class);

    public static List<File> unzip(String decompressingFilesDirectory, String zipFile) {
        try {
            logger.info("文件信息：{} \\ {}", decompressingFilesDirectory, zipFile);
            return unzipFile(decompressingFilesDirectory, zipFile);
        } catch (FileNotFoundException fNFEx) {
            logger.error("解压ZIP文件发生异常：文件不存在!", fNFEx);
        } catch (IOException ioEx) {
            logger.error("解压ZIP文件发生异常：IO异常!", ioEx);
        } catch (Exception ex) {
            logger.error("解压ZIP文件发生系统异常!", ex);
        } catch (Throwable t) {
            logger.error("解压ZIP文件发生系统底层异常!", t);
        }
        return null;
    }

    private static List<File> unzipFile(String decompressingFilesDirectory, String zipFile)
                                                                                           throws FileNotFoundException,
                                                                                           IOException {
        ZipInputStream zipInputStream = new ZipInputStream(new FileInputStream(zipFile));
        BufferedInputStream bufferedInputStream = new BufferedInputStream(zipInputStream);
        List<File> files = new ArrayList<File>();
        File file = null;
        ZipEntry zipEntry;
        while ((zipEntry = zipInputStream.getNextEntry()) != null) {
            if (zipEntry.isDirectory()) {
                continue;
            }

            file = new File(decompressingFilesDirectory, zipEntry.getName());
            if (!file.exists()) {
                (new File(file.getParent())).mkdirs();
            }

            FileOutputStream fileOutPutStream = new FileOutputStream(file);
            BufferedOutputStream bufferedOutPutStream = new BufferedOutputStream(fileOutPutStream);
            int b;
            while ((b = bufferedInputStream.read()) != -1) {
                bufferedOutPutStream.write(b);
            }
            bufferedOutPutStream.close();
            fileOutPutStream.close();
            files.add(file);
        }
        bufferedInputStream.close();
        zipInputStream.close();

        return files;
    }
}
