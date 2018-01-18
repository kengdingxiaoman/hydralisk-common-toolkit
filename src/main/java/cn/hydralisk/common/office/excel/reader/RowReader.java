package cn.hydralisk.common.office.excel.reader;

import java.util.List;

/**
 * created by master.yang 2018/1/18 下午3:19
 */
public interface RowReader {

    void readRow(List<String> cellValues);
}
