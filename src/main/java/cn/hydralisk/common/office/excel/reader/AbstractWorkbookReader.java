package cn.hydralisk.common.office.excel.reader;

import org.apache.poi.ss.usermodel.*;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 使用Workbook方式读取excel的抽象方法
 * created by master.yang 2017-12-14 下午5:44
 */
public abstract class AbstractWorkbookReader implements IExcelReader{

    private InputStream inputStream;
    private int startRow;
    private int endRowIgnoredNum;

    public AbstractWorkbookReader(InputStream inputStream, int startRow, int endRowIgnoredNum) {
        this.inputStream = inputStream;
        this.startRow = startRow;
        this.endRowIgnoredNum = endRowIgnoredNum;
    }

    public void read(RowReader reader) throws ExcelReadException{
        Workbook workbook;
        try {
            workbook = initWorkBook();
        } catch(Exception ex) {
            throw new ExcelReadException("初始化workbook发生异常", ex);
        }
        int sheetNum = workbook.getNumberOfSheets();
        for (int i = 0; i < sheetNum; i++) {
            readSheet(workbook.getSheetAt(i), reader);
        }
    }

    protected abstract Workbook initWorkBook() throws Exception;

    private void readSheet(Sheet sheet, RowReader reader) {
        int lastRowNum = sheet.getLastRowNum();
        int ignoredRowStartNum = lastRowNum - this.endRowIgnoredNum;

        for (int rowNum = sheet.getFirstRowNum(); rowNum <= lastRowNum; rowNum++) {
            if (rowNum < startRow - 1) {
                continue;
            }
            if (rowNum > ignoredRowStartNum) {
                break;
            }

            List<String> rowValues = readRow(sheet.getRow(rowNum));
            reader.readRow(rowValues);
        }
    }

    private List<String> readRow(Row row) {
        List<String> data = new ArrayList<>();

        int lastCellNum = row.getLastCellNum();
        for (int cellNum = row.getFirstCellNum(); cellNum < lastCellNum; cellNum++) {
            Cell cell = row.getCell(cellNum);
            Object o = getCellValue(cell);
            data.add(o == null ? null : o.toString());
        }

        return data;
    }

    private Object getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        CellType cellType = cell.getCachedFormulaResultTypeEnum();

        switch(cellType) {
            case BLANK:
            case STRING:
                return cell.getRichStringCellValue().getString();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return CellValueConvertUtils.format(cell.getDateCellValue());
                } else {
                    return CellValueConvertUtils.convertToNumberSmart(cell.getNumericCellValue());
                }
            case FORMULA: return CellValueConvertUtils.convertFormulaCellValue(cell);
            case BOOLEAN: return cell.getBooleanCellValue();
            case ERROR: return cell.getErrorCellValue();
            default: return null;
        }
    }

    //getter
    protected InputStream getInputStream(){
        return inputStream;
    }
}
