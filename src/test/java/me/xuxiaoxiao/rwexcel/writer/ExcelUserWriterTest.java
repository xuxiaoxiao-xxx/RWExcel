package me.xuxiaoxiao.rwexcel.writer;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.reader.ExcelReader;
import me.xuxiaoxiao.rwexcel.reader.ExcelStreamReader;
import me.xuxiaoxiao.rwexcel.reader.ExcelStreamReaderTest;
import org.junit.Test;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.LinkedList;
import java.util.List;

public class ExcelUserWriterTest {

    @Test
    public void writeXls() throws Exception {
        ByteArrayOutputStream baOutStream = new ByteArrayOutputStream();

        ExcelWriter excelWriter = new ExcelUserWriter();
        excelWriter.write(baOutStream, new TestProvider(ExcelWriter.Version.XLS));

        ExcelReader excelReader = new ExcelStreamReader();
        excelReader.read(new ByteArrayInputStream(baOutStream.toByteArray()), new ExcelStreamReaderTest.TestListener());
    }

    public static final class TestProvider implements ExcelWriter.Provider {
        private ExcelWriter.Version version;

        public TestProvider(@Nonnull ExcelWriter.Version version) {
            this.version = version;
        }

        @Nonnull
        @Override
        public ExcelWriter.Version version() {
            return this.version;
        }

        @Nullable
        @Override
        public ExcelSheet provideSheet(int lastSheetIndex) {
            if (lastSheetIndex == -1) {
                return new ExcelSheet(0, "Sheet1");
            } else if (lastSheetIndex == 0) {
                return new ExcelSheet(1, "Sheet2");
            } else {
                return null;
            }
        }

        @Nullable
        @Override
        public ExcelRow provideRow(@Nonnull ExcelSheet sheet, int lastRowIndex) {
            if (sheet.getShtIndex() == 0 && lastRowIndex < 12) {
                ExcelRow row = new ExcelRow(sheet.getShtIndex(), lastRowIndex + 1);
                if (lastRowIndex < 6) {
                    row.setColFirst(0);
                    row.setColLast(row.getRowIndex());
                } else {
                    row.setColFirst(row.getRowIndex() - 6);
                    row.setColLast(6);
                }
                System.out.println("创建row：shtIndex=" + sheet.getShtIndex() + ",rowIndex=" + row.getRowIndex() + "，colFirst=" + row.getColFirst() + "，colLast=" + row.getColLast());
                return row;
            } else if (sheet.getShtIndex() == 1 && lastRowIndex < 4) {
                ExcelRow row = new ExcelRow(sheet.getShtIndex(), lastRowIndex + 1);
                if (lastRowIndex < 2) {
                    row.setColFirst(0);
                    row.setColLast(row.getRowIndex());
                } else {
                    row.setColFirst(row.getRowIndex() - 2);
                    row.setColLast(2);
                }
                System.out.println("创建row：shtIndex=" + sheet.getShtIndex() + ",rowIndex=" + row.getRowIndex() + "，colFirst=" + row.getColFirst() + "，colLast=" + row.getColLast());
                return row;
            } else {
                return null;
            }
        }

        @Nonnull
        @Override
        public List<ExcelCell> provideCells(@Nonnull ExcelSheet sheet, @Nonnull ExcelRow row) {
            List<ExcelCell> cells = new LinkedList<>();
            if (sheet.getShtIndex() == 0) {
                for (int i = row.getColFirst(); i <= row.getColLast(); i++) {
                    ExcelCell cell = new ExcelCell(sheet.getShtIndex(), row.getRowIndex(), i, (char) ('A' + i) + "" + (row.getRowIndex() + 1));
                    System.out.println("colIndex=" + cell.getColIndex() + "，strValue=" + cell.getStrValue());
                    cells.add(cell);
                }
            } else {
                for (int i = row.getColFirst(); i <= row.getColLast(); i++) {
                    ExcelCell cell = new ExcelCell(sheet.getShtIndex(), row.getRowIndex(), i, (row.getRowIndex() + 1) + "" + (char) ('A' + i));
                    System.out.println("colIndex=" + cell.getColIndex() + "，strValue=" + cell.getStrValue());
                    cells.add(cell);
                }
            }
            return cells;
        }
    }
}