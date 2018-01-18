package cn.hydralisk.common.office.excel.reader;

import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * HSSF大文件读取
 * created by master.yang 2017-12-14 上午11:32
 */
public class HSSFEventUserModelReader implements IExcelReader, HSSFListener {

    private InputStream inputStream;
    private int startRow;
    private int endRowIgnoredNum;

    private int minColumns;

    private int lastRowNumber; //上一次的行号，该值会随时变化
    private int lastColumnNumber; //上一次的列号，该值会随时变化

    private int sheetLastRowNum; //sheet的最后一行，该值不变

    /** Should we output the formula, or the value it has? */
    private boolean outputFormulaValues = true;

    /** For parsing Formulas */
    private EventWorkbookBuilder.SheetRecordCollectingListener workbookBuildingListener;
    private HSSFWorkbook stubWorkbook;

    // Records we pick up as we process
    private SSTRecord sstRecord;
    private FormatTrackingHSSFListener formatListener;

    /** So we known which sheet we're on */
    private int sheetIndex = -1;
    private BoundSheetRecord[] orderedBSRs;
    @SuppressWarnings("unchecked")
    private ArrayList boundSheetRecords = new ArrayList();

    // For handling formulas with string results
    private int nextRow;
    private int nextColumn;
    private boolean outputNextStringRecord;

    private int curRow;
    private List<String> rowlist;
    @SuppressWarnings( "unused")
    private String sheetName;

    private RowReader rowReader;

    public HSSFEventUserModelReader(InputStream inputStream, int startRow, int endRowIgnoredNum) {
        this.inputStream = inputStream;
        this.startRow = startRow;
        this.endRowIgnoredNum = endRowIgnoredNum;

        this.minColumns = -1;
        this.curRow = 0;
        this.rowlist = new ArrayList<String>();
    }

    @Override
    public void read(RowReader reader) throws ExcelReadException{
        this.rowReader = reader;

        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
        formatListener = new FormatTrackingHSSFListener(listener);

        HSSFEventFactory factory = new HSSFEventFactory();
        HSSFRequest request = new HSSFRequest();

        if(outputFormulaValues) {
            request.addListenerForAllRecords(formatListener);
        } else {
            workbookBuildingListener = new EventWorkbookBuilder.SheetRecordCollectingListener(formatListener);
            request.addListenerForAllRecords(workbookBuildingListener);
        }

        try{
            POIFSFileSystem poifsFileSystem = new POIFSFileSystem(inputStream);
            factory.processWorkbookEvents(request, poifsFileSystem);
        }catch(IOException ioEx){
            throw new ExcelReadException("HSSFEvent模式读取excel文件发生异常", ioEx);
        }
    }

    /**
     * Main HSSFListener method, processes events
     */
    public void processRecord(Record record) {
        int thisRow = -1;
        int thisColumn = -1;
        String thisStr = null;
        String value = null;

        if (record instanceof CellValueRecordInterface) {
            thisRow = ((CellValueRecordInterface)record).getRow();
            thisColumn = ((CellValueRecordInterface)record).getColumn();

            if (thisRow > -1) {

                if (ignoreThisRow(thisRow)) {

                    if (thisRow > -1) {
                        lastRowNumber = thisRow;
                    }
                    if (thisColumn > -1) {
                        lastColumnNumber = thisColumn;
                    }
                    return ;
                }
            }
        }

        if (record instanceof MissingCellDummyRecord) {
            thisRow = ((MissingCellDummyRecord)record).getRow();
            thisColumn = ((MissingCellDummyRecord)record).getColumn();

            if (thisRow > -1) {

                if (ignoreThisRow(thisRow)) {

                    if (thisRow > -1) {
                        lastRowNumber = thisRow;
                    }
                    if (thisColumn > -1) {
                        lastColumnNumber = thisColumn;
                    }
                    return ;
                }
            }
        }

        switch (record.getSid()) {
            case BoundSheetRecord.sid:
                boundSheetRecords.add(record);
                break;
            case DimensionsRecord.sid:
                DimensionsRecord dimensions = (DimensionsRecord) record;
                sheetLastRowNum = dimensions.getLastRow();
                break;
            case BOFRecord.sid:
                BOFRecord br = (BOFRecord) record;
                if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
                    // Create sub workbook if required
                    if (workbookBuildingListener != null && stubWorkbook == null) {
                        stubWorkbook = workbookBuildingListener
                                .getStubHSSFWorkbook();
                    }

                    // Works by ordering the BSRs by the location of
                    // their BOFRecords, and then knowing that we
                    // process BOFRecords in byte offset order
                    sheetIndex++;
                    if (orderedBSRs == null) {
                        orderedBSRs = BoundSheetRecord
                                .orderByBofPosition(boundSheetRecords);
                    }
                    sheetName = orderedBSRs[sheetIndex].getSheetname();
                }
                break;

            case SSTRecord.sid:
                sstRecord = (SSTRecord) record;
                break;

            case BlankRecord.sid:
                BlankRecord brec = (BlankRecord) record;

                thisRow = brec.getRow();
                thisColumn = brec.getColumn();
                thisStr = "";
                rowlist.add(thisColumn, "");
                break;
            case BoolErrRecord.sid:
                BoolErrRecord berec = (BoolErrRecord) record;

                thisRow = berec.getRow();
                thisColumn = berec.getColumn();
                thisStr = "";
                rowlist.add(thisColumn, (new Boolean(berec.getBooleanValue())).toString());
                break;

            case FormulaRecord.sid:
                FormulaRecord frec = (FormulaRecord) record;

                thisRow = frec.getRow();
                thisColumn = frec.getColumn();

                if (outputFormulaValues) {
                    if (Double.isNaN(frec.getValue())) {
                        // Formula result is a string
                        // This is stored in the next record
                        outputNextStringRecord = true;
                        nextRow = frec.getRow();
                        nextColumn = frec.getColumn();
                    } else {
                        thisStr = CellValueConvertUtils.convertToNumberSmart(frec.getValue());
                    }
                } else {
                    thisStr = '"' + HSSFFormulaParser.toFormulaString(stubWorkbook,
                            frec.getParsedExpression()) + '"';
                }
                rowlist.add(thisColumn, thisStr);
                break;
            case StringRecord.sid:
                if (outputNextStringRecord) {
                    // String for formula
                    StringRecord srec = (StringRecord) record;
                    thisStr = srec.getString();
                    thisRow = nextRow;
                    thisColumn = nextColumn;
                    outputNextStringRecord = false;
                    rowlist.add(thisColumn, thisStr);
                }
                break;

            case LabelRecord.sid:
                LabelRecord lrec = (LabelRecord) record;

                curRow = thisRow = lrec.getRow();
                thisColumn = lrec.getColumn();
                value = lrec.getValue().trim();
                value = value.equals("")?" ":value;
                this.rowlist.add(thisColumn, value);
                break;
            case LabelSSTRecord.sid:
                LabelSSTRecord lsrec = (LabelSSTRecord) record;

                curRow = thisRow = lsrec.getRow();
                thisColumn = lsrec.getColumn();
                if (sstRecord == null) {
                    rowlist.add(thisColumn, " ");
                } else {
                    value =  sstRecord
                            .getString(lsrec.getSSTIndex()).toString().trim();
                    value = value.equals("")?" ":value;
                    rowlist.add(thisColumn, value);
                }
                break;
            case NoteRecord.sid:
                NoteRecord nrec = (NoteRecord) record;

                thisRow = nrec.getRow();
                thisColumn = nrec.getColumn();
                // TODO: Find object to match nrec.getShapeId()
                thisStr = '"' + "(TODO)" + '"';
                rowlist.add(thisColumn, thisStr);
                break;
            case NumberRecord.sid:
                NumberRecord numrec = (NumberRecord) record;

                curRow = thisRow = numrec.getRow();
                thisColumn = numrec.getColumn();
                value = formatListener.formatNumberDateCell(numrec).trim();
                value = value.equals("")?" ":value;
                // Format
                rowlist.add(thisColumn, CellValueConvertUtils.convertToNumberSmart(Double.valueOf(value)));
                break;
            case RKRecord.sid:
                RKRecord rkrec = (RKRecord) record;

                thisRow = rkrec.getRow();
                thisColumn = rkrec.getColumn();
                thisStr = '"' + "(TODO)" + '"';
                rowlist.add(thisColumn, value);
                break;
            default:
                break;
        }

        // 遇到新行的操作
        if (thisRow != -1 && thisRow != lastRowNumber) {
            lastColumnNumber = -1;
        }

        // 空值的操作
        if (record instanceof MissingCellDummyRecord) {
            MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
            curRow = thisRow = mc.getRow();
            thisColumn = mc.getColumn();
            rowlist.add(thisColumn," ");
        }

        // 更新行和列的值
        if (thisRow > -1)
            lastRowNumber = thisRow;
        if (thisColumn > -1)
            lastColumnNumber = thisColumn;

        // 行结束时的操作
        if (record instanceof LastCellOfRowDummyRecord) {
            if (minColumns > 0) {
                // 列值重新置空
                if (lastColumnNumber == -1) {
                    lastColumnNumber = 0;
                }
            }
            lastColumnNumber = -1;

            if (! rowlist.isEmpty()) {
                rowReader.readRow(rowlist);
                rowlist.clear();
            }
        }
    }

    /**
     * 如果设置了跳过头和尾，那就忽略该行
     * @param thisRow
     * @return
     */
    private boolean ignoreThisRow(int thisRow) {
        if (startRow > 0 && thisRow < (startRow - 1)) {
            return true;
        }
        if (endRowIgnoredNum > 0 && thisRow > (sheetLastRowNum - endRowIgnoredNum - 1)) {
            return true;
        }

        return false;
    }

    @Override
    public void close() throws IOException{
        if(this.inputStream != null) {
            inputStream.close();
        }
    }
}
