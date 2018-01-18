package cn.hydralisk.common.office.excel.reader;

import org.apache.poi.poifs.filesystem.FileMagic;

import java.io.IOException;
import java.io.InputStream;

/**
 * 返回合适的读取excel实现类
 * created by master.yang 2017-12-14 下午5:31
 */
public class ExcelReaderFactory {

    public static IExcelReader generateReader(InputStream inputStream, int startRow, int endRowIgnoredNum) throws ExcelReadException{
        try{
            return generate(inputStream, startRow, endRowIgnoredNum);
        }catch(Exception ioEx) {
            throw new ExcelReadException("分析文件发生IO异常", ioEx);
        }
    }

    /**
     * 生成具体的实际读取文件类
     * @param inputStream
     * @param startRow 从第几行开始读取
     * @param endRowIgnoredNum 忽略末尾多少行数
     * @return 具体的实现类
     * @throws IOException
     */
    private static IExcelReader generate(InputStream inputStream, int startRow, int endRowIgnoredNum) throws IOException{
        if (inputStream == null) {
            throw new IllegalArgumentException("inputStream is null");
        }

        /**
         * 判断excel文件版本
         * @see org.apache.poi.ss.usermodel.WorkbookFactory#create(InputStream)
         */
        InputStream is = FileMagic.prepareToCheckMagic(inputStream);
        FileMagic fileMagic = FileMagic.valueOf(is);

        //根据不同的文件类型(xls, xlsx)使用不同的处理类
        if (FileMagic.OLE2.equals(fileMagic)) {
            return initHSSFReader(is, startRow, endRowIgnoredNum);
        }
        if (FileMagic.OOXML.equals(fileMagic)) {
            return initXSSFReader(is, startRow, endRowIgnoredNum);
        }

        throw new IllegalArgumentException("Your InputStream was neither an OLE2 stream, nor an OOXML stream");
    }

    private static IExcelReader initXSSFReader(InputStream inputStream, int startRow, int endRowIgnoredNum) throws IOException{
        if (ExcelUtils.isLargeFile(inputStream)) {
            return new XSSFSaxReader(inputStream, startRow, endRowIgnoredNum);
        }

        return new XSSFWorkbookReader(inputStream, startRow, endRowIgnoredNum);
    }

    private static IExcelReader initHSSFReader(InputStream inputStream, int startRow, int endRowIgnoredNum) throws IOException{
        if (ExcelUtils.isLargeFile(inputStream)) {
            return new HSSFEventUserModelReader(inputStream, startRow, endRowIgnoredNum);
        }

        return new HSSFWorkbookReader(inputStream, startRow, endRowIgnoredNum);
    }
}
