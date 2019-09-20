package me.xuxiaoxiao.rwexcel.writer;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import org.junit.Test;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.io.FileOutputStream;
import java.util.Collections;
import java.util.List;

public class ExcelWriterImplTest {

    @Test
    public void write() throws Exception {
        ExcelWriter excelWriter = new ExcelWriterImpl();
        excelWriter.write(new FileOutputStream("test.xls"), new ExcelWriter.Provider() {
            @Nonnull
            @Override
            public ExcelWriter.Type version() {
                return ExcelWriter.Type.XLS;
            }

            @Nonnull
            @Override
            public ExcelSheet[] sheets() {
                return new ExcelSheet[]{new ExcelSheet(0, "测试sheet")};
            }

            @Nullable
            @Override
            public ExcelRow provideRow(ExcelSheet sheet, int lastRowIndex) {
                if (lastRowIndex == -1) {
                    ExcelRow excelRow = new ExcelRow(sheet.getShtIndex(), 0);
                    excelRow.setColFirst(1);
                    excelRow.setColLast(3);
                    return excelRow;
                } else if (lastRowIndex == 0) {
                    ExcelRow excelRow = new ExcelRow(sheet.getShtIndex(), 3);
                    excelRow.setColFirst(1);
                    excelRow.setColLast(4);
                    return excelRow;
                } else {
                    return null;
                }
            }

            @Nonnull
            @Override
            public List<ExcelCell> provideCells(ExcelSheet sheet, ExcelRow row) {
                return Collections.singletonList(new ExcelCell(sheet.getShtIndex(), row.getRowIndex(), 0, "row=" + row.getRowIndex() + ",col=0"));
            }
        });
    }
}