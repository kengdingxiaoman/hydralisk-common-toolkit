package cn.hydralisk.common.office.excel.reader;

import org.apache.poi.ss.usermodel.DateUtil;

import java.io.IOException;
import java.io.InputStream;
import java.util.Date;

/**
 * 工具类，提供一些方法
 * created by master.yang 2017-12-18 下午1:56
 */
public abstract class ExcelUtils {

    private static final int LARGE_FILE_THRESHOLD = 1024 * 1024 * 2; //2M以上视为大文件

    /**
     * 判断是否是大文件，大文件读取时会使用不同的处理实现
     * inputStream.available() 因为受到 Integer.MAX_VALUE 值的影响，最多返回2G
     * @param inputStream
     * @return boolean 是否是大文件
     * @throws IOException
     */
    public static boolean isLargeFile(InputStream inputStream) throws IOException{
        if (inputStream == null) {
            throw new IllegalArgumentException("inputStream is null");
        }

        int fileSize = inputStream.available(); //这样获取的fileSize不精确，但对于判断已经够了

        return isLargeFile(fileSize);
    }

    public static boolean isLargeFile(int fileSize) throws IOException{
        if (fileSize <= 0) {
            throw new IllegalArgumentException("ile size is illegal, file maybe has not be read.");
        }

        return fileSize > LARGE_FILE_THRESHOLD;
    }

    /**
     * 将excelDate转换成javaDate
     * excelDate: 43034.484143518515
     * javaDate: 2017-10-26 11:37:10
     * @param str
     * @return
     */
    public static final String convertToJavaDate(String str){
        Date javaDate= DateUtil.getJavaDate(Double.valueOf(str));
        return CellValueConvertUtils.format(javaDate);
    }

    /**
     * 有的日期会解析成为excel date，形似：43034.484143518515 或 43056
     * @param dateString
     * @return
     */
    public static boolean isExcelDateFormat(String dateString) {
        return dateString.matches("[\\d]+\\.[\\d]+") || dateString.matches("[\\d]+");
    }
}
