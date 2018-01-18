package cn.hydralisk.common.office.excel.reader;

import cn.hydralisk.common.exception.HydraliskException;

/**
 * Excel读取异常
 * @author master.yang
 */
public class ExcelReadException extends HydraliskException {

	private static final long serialVersionUID = 1L;

	public ExcelReadException(String message) {
		super(message);
	}
	
	public ExcelReadException(String message, Throwable e) {
		super(message, e);
	}
}
