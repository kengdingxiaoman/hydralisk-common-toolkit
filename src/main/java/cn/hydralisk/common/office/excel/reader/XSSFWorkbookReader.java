package cn.hydralisk.common.office.excel.reader;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;

/**
 * XSSF读取xlxs
 * @author master.yang
 */
public class XSSFWorkbookReader extends AbstractWorkbookReader{

	private OPCPackage opcPackage;

	public XSSFWorkbookReader(InputStream inputStream, int startRow, int endRowIgnoredNum) {
		super(inputStream, startRow, endRowIgnoredNum);
	}

	@Override
	protected Workbook initWorkBook() throws Exception{
		this.opcPackage = OPCPackage.open(this.getInputStream());
		return new XSSFWorkbook(opcPackage);
	}

	@Override
	public void close() throws IOException{
		if (opcPackage != null) {
			opcPackage.close();
		}
	}
}
