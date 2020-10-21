package me.xuxiaoxiao.rwexcel.reader;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.record.*;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import javax.annotation.Nonnull;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

/**
 * Excel解析器EventModel实现
 * <ul>
 * <li>[2019/9/12 19:42]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public class ExcelStreamReader implements ExcelReader {

    @Override
    public void read(@Nonnull InputStream inStream, @Nonnull Listener listener) throws Exception {
        inStream = FileMagic.prepareToCheckMagic(inStream);
        FileMagic magic = FileMagic.valueOf(inStream);
        switch (magic) {
            case OLE2:
                try (POIFSFileSystem pStream = new POIFSFileSystem(inStream)) {
                    XlsScanner xlsScanner = new XlsScanner(listener);
                    xlsScanner.setFormatListener(new FormatTrackingHSSFListener(new MissingRecordAwareHSSFListener(xlsScanner)));

                    HSSFRequest hssfRequest = new HSSFRequest();
                    hssfRequest.addListenerForAllRecords(xlsScanner.getFormatListener());

                    HSSFEventFactory eventFactory = new HSSFEventFactory();
                    eventFactory.processWorkbookEvents(hssfRequest, pStream);
                }
                break;
            case OOXML:
                XSSFReader reader = new XSSFReader(OPCPackage.open(inStream));
                XSSFReader.SheetIterator iterator = (XSSFReader.SheetIterator) reader.getSheetsData();
                int sheetIndex = -1;
                while (iterator.hasNext()) {
                    try (InputStream stream = iterator.next()) {
                        ExcelSheet sheet = new ExcelSheet(++sheetIndex, iterator.getSheetName());
                        listener.onSheetStart(sheet);

                        XMLReader sheetParser = SAXHelper.newXMLReader();
                        sheetParser.setContentHandler(new XSSFSheetXMLHandler(reader.getStylesTable(), reader.getSharedStringsTable(), new XlsxScanner(sheet, listener), new ExcelReader.Formatter(), false));
                        sheetParser.parse(new InputSource(stream));
                    }
                }
                break;
            default:
                throw new IllegalArgumentException("未能识别Excel文件");
        }
        inStream.close();
    }

    private static class XlsScanner implements HSSFListener {
        private final SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss.SSS");
        private Listener listener;
        private FormatTrackingHSSFListener formatListener;

        /**
         * 字符串缓存
         */
        private SSTRecord cacheString;

        /**
         * 当前sheet
         */
        private ExcelSheet sheet;
        private BoundSheetRecord[] orderedBSRs;
        private List<BoundSheetRecord> boundSheetRecords = new ArrayList<>();

        /**
         * 当前行
         */
        private int rowIndex = -1;
        private int colFirst = -1, colLast = -1;
        private List<String> valList = new LinkedList<>();

        /**
         * 字符串公式相关
         */
        private int nextStringRecordColumn = -1;
        private boolean handleNextStringRecord = false;

        public XlsScanner(Listener listener) {
            this.listener = listener;
        }

        public FormatTrackingHSSFListener getFormatListener() {
            return formatListener;
        }

        public void setFormatListener(FormatTrackingHSSFListener formatListener) {
            this.formatListener = formatListener;
        }

        public void finishRow() {
            if (rowIndex >= 0 && colFirst >= 0 && colLast >= colFirst) {
                //该行存在并且有数据
                for (int i = 0; i <= colLast - colFirst; i++) {
                    //去除开头的空白单元格，重新计算colFirst
                    if (valList.get(i) == null) {
                        valList.remove(i--);
                        colFirst++;
                    } else {
                        break;
                    }
                }
                for (int i = colLast - colFirst; i >= 0; i--) {
                    //去除结尾的空白单元格，重新计算colLast
                    if (valList.get(i) == null) {
                        valList.remove(i);
                        colLast--;
                    } else {
                        break;
                    }
                }
                if (colFirst >= 0 && colLast >= colFirst) {
                    //仍然有非空单元格
                    ExcelRow excelRow = new ExcelRow(this.sheet.getShtIndex(), rowIndex);
                    excelRow.setColFirst(colFirst);
                    excelRow.setColLast(colLast);
                    List<ExcelCell> cells = new LinkedList<>();
                    for (int i = colFirst; i <= colLast; i++) {
                        if (valList.get(i - colFirst) != null) {
                            cells.add(new ExcelCell(this.sheet.getShtIndex(), rowIndex, i, valList.get(i - colFirst)));
                        }
                    }
                    listener.onRow(this.sheet, excelRow, cells);
                }
            }
            rowIndex = -1;
            colFirst = -1;
            colLast = -1;
            valList.clear();
        }

        @Override
        public void processRecord(Record record) {
            if (record instanceof CellValueRecordInterface) {
                CellValueRecordInterface cellRecord = (CellValueRecordInterface) record;
                if (cellRecord.getRow() != rowIndex) {
                    rowIndex = cellRecord.getRow();
                    colFirst = cellRecord.getColumn();
                    colLast = cellRecord.getColumn();
                } else {
                    if (colFirst < 0) {
                        colFirst = cellRecord.getColumn();
                        colLast = cellRecord.getColumn();
                    } else {
                        while (++colLast < cellRecord.getColumn()) {
                            valList.add(null);
                        }
                    }
                }

                switch (record.getSid()) {
                    case BlankRecord.sid:
                        //空白单元格，不处理
                        valList.add(null);
                        break;
                    case BoolErrRecord.sid:
                        BoolErrRecord boolErrRecord = (BoolErrRecord) record;
                        valList.add(String.valueOf(!boolErrRecord.isError() && boolErrRecord.getBooleanValue()));
                        break;
                    case FormulaRecord.sid:
                        FormulaRecord formulaRecord = (FormulaRecord) record;
                        switch (formulaRecord.getCachedResultType()) {
                            case 0: {
                                //数值型
                                valList.add(formatListener.formatNumberDateCell(formulaRecord));
                                break;
                            }
                            case 1: {
                                //字符串单元格
                                handleNextStringRecord = true;
                                nextStringRecordColumn = formulaRecord.getColumn();
                                valList.add(null);
                                break;
                            }
                            case 3: {
                                //空白单元格
                                valList.add("");
                                break;
                            }
                            case 4: {
                                //逻辑值单元格
                                valList.add(String.valueOf(formulaRecord.getCachedBooleanValue()));
                                break;
                            }
                            default: {
                                //未知或错误单元格，添加null占位，最终会当成空白单元格处理
                                valList.add(null);
                                break;
                            }
                        }
                        break;
                    case LabelRecord.sid:
                        LabelRecord labelRecord = (LabelRecord) record;
                        valList.add(labelRecord.getValue());
                        break;
                    case LabelSSTRecord.sid:
                        LabelSSTRecord labelSSTRecord = (LabelSSTRecord) record;
                        valList.add(cacheString.getString(labelSSTRecord.getSSTIndex()).getString());
                        break;
                    case NumberRecord.sid:
                        NumberRecord numberRecord = (NumberRecord) record;
                        String formatValue = formatListener.formatNumberDateCell(numberRecord);

                        int formatIndex = formatListener.getFormatIndex(numberRecord);
                        String formatString = formatListener.getFormatString(numberRecord);
                        if (DateUtil.isADateFormat(formatIndex, formatString)) {
                            if (DateUtil.isValidExcelDate(numberRecord.getValue())) {
                                //是日期类型值
                                Date date = DateUtil.getJavaDate(numberRecord.getValue(), false);
                                formatValue = sdf.format(date);
                            }
                        }
                        valList.add(formatValue);
                        break;
                    case RKRecord.sid:
                        RKRecord rkRecord = (RKRecord) record;
                        valList.add(String.valueOf(rkRecord.getRKNumber()));
                        break;
                    default:
                        System.out.println("未识别：" + record.getClass().getSimpleName());
                }
            } else {
                switch (record.getSid()) {
                    case BoundSheetRecord.sid:
                        boundSheetRecords.add((BoundSheetRecord) record);
                        break;
                    case BOFRecord.sid:
                        BOFRecord bofRecord = (BOFRecord) record;
                        if (bofRecord.getType() == BOFRecord.TYPE_WORKSHEET) {
                            if (orderedBSRs == null) {
                                orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
                            }
                            if (this.sheet == null) {
                                this.sheet = new ExcelSheet(0, orderedBSRs[0].getSheetname());
                            } else {
                                this.sheet = new ExcelSheet(this.sheet.getShtIndex() + 1, orderedBSRs[this.sheet.getShtIndex() + 1].getSheetname());
                            }
                            rowIndex = -1;
                            colFirst = -1;
                            colLast = -1;
                            listener.onSheetStart(this.sheet);
                        }
                        break;
                    case SSTRecord.sid:
                        cacheString = (SSTRecord) record;
                        break;
                    case StringRecord.sid:
                        if (handleNextStringRecord) {
                            StringRecord stringRecord = (StringRecord) record;
                            valList.set(nextStringRecordColumn - colFirst, stringRecord.getString());
                            handleNextStringRecord = false;
                        }
                        break;
                    case EOFRecord.sid:
                        if (this.sheet != null) {
                            listener.onSheetEnd(this.sheet);
                        }
                        break;
                }
            }
            if (record instanceof LastCellOfRowDummyRecord) {
                finishRow();
            }
        }
    }

    public static class XlsxScanner implements XSSFSheetXMLHandler.SheetContentsHandler {
        private final ExcelSheet sheet;
        private final Listener listener;

        private int colFirst = -1, colLast = -1;
        private List<String> valList = new LinkedList<>();

        public XlsxScanner(ExcelSheet sheet, Listener listener) {
            this.sheet = sheet;
            this.listener = listener;
        }

        @Override
        public void startRow(int rowNum) {
        }

        @Override
        public void endRow(int rowNum) {
            if (colFirst >= 0 && colLast >= colFirst) {
                ExcelRow excelRow = new ExcelRow(this.sheet.getShtIndex(), rowNum);
                excelRow.setColFirst(colFirst);
                excelRow.setColLast(colLast);
                List<ExcelCell> cells = new LinkedList<>();
                for (int i = colFirst; i <= colLast; i++) {
                    if (valList.get(i - colFirst) != null) {
                        cells.add(new ExcelCell(this.sheet.getShtIndex(), rowNum, i, valList.get(i - colFirst)));
                    }
                }
                listener.onRow(this.sheet, excelRow, cells);
            }
            colFirst = -1;
            colLast = -1;
            valList.clear();
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment xssfComment) {
            int col = new CellReference(cellReference).getCol();
            if (colFirst < 0) {
                colFirst = col;
                colLast = col;
                valList.add(formattedValue);
            } else {
                while (++colLast < col) {
                    valList.add(null);
                }
                valList.add(formattedValue);
            }
        }

        @Override
        public void headerFooter(String text, boolean isHeader, String tagName) {
        }

        @Override
        public void endSheet() {
            listener.onSheetEnd(this.sheet);
        }
    }
}
