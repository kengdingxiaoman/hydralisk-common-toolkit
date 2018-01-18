package cn.hydralisk.common.office.excel.reader;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.InputStream;

/**
 * HSSF读取xls文件
 * created by master.yang 2017-12-14 上午11:31
 */
public class HSSFWorkbookReader extends AbstractWorkbookReader{

    private NPOIFSFileSystem npoifsFileSystem;

    public HSSFWorkbookReader(InputStream inputStream, int startRow, int endRowIgnoredNum) {
        super(inputStream, startRow, endRowIgnoredNum);
    }

    @Override
    protected Workbook initWorkBook() throws Exception{
        this.npoifsFileSystem = new NPOIFSFileSystem(this.getInputStream());
        return new HSSFWorkbook(npoifsFileSystem.getRoot(), true);
    }

    @Override
    public void close() throws IOException{
        if(npoifsFileSystem != null) {
            this.npoifsFileSystem.close();
        }
    }
}
