package cn.hydralisk.common.office.excel.reader;

import java.io.Closeable;

/**
 * 定义方法
 * created by master.yang 2017-12-14 下午5:29
 */
public interface IExcelReader extends Closeable{

    void read(RowReader reader) throws ExcelReadException;
}
