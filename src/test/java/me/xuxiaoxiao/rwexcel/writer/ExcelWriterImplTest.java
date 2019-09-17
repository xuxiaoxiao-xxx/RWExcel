package me.xuxiaoxiao.rwexcel.writer;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import org.junit.Test;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.io.FileOutputStream;

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

            @Nullable
            @Override
            public ExcelSheet nextSheet(int lastShtIndex) {
                if (lastShtIndex == -1) {
                    return new ExcelSheet(0, "测试sheet");
                } else {
                    return null;
                }
            }

            @Nullable
            @Override
            public ExcelRow nextRow(int shtIndex, int lastRowIndex) {
                if (lastRowIndex == -1) {
                    ExcelRow excelRow = new ExcelRow(shtIndex, 0);
                    excelRow.setColFirst(1);
                    excelRow.setColLast(3);
                    return excelRow;
                } else if (lastRowIndex == 0) {
                    ExcelRow excelRow = new ExcelRow(shtIndex, 3);
                    excelRow.setColFirst(1);
                    excelRow.setColLast(4);
                    return excelRow;
                } else {
                    return null;
                }
            }

            @Nullable
            @Override
            public ExcelCell provideCell(int shtIndex, int rowIndex, int colIndex) {
                return new ExcelCell(shtIndex, rowIndex, colIndex, "row=" + rowIndex + ",col=" + colIndex);
            }
        });

    }
}