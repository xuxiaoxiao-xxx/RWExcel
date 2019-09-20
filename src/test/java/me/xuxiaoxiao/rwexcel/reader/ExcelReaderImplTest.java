package me.xuxiaoxiao.rwexcel.reader;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import org.junit.Test;

import javax.annotation.Nonnull;
import java.util.List;
import java.util.Objects;

public class ExcelReaderImplTest {

    @Test
    public void read() throws Exception {
        ExcelReader reader = new ExcelReaderImpl();
        reader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("test.xls")), new ExcelReader.Listener() {

            @Override
            public void onSheetStart(@Nonnull ExcelSheet sheet) {
                System.out.println();
                System.out.println("开始sheet：" + sheet.getShtName());
            }

            @Override
            public void onRow(@Nonnull ExcelRow row, @Nonnull List<ExcelCell> cells) {
                System.out.println("获取到row：" + row.getRowIndex() + " ，firstCol=" + row.getColFirst() + " ，lastCol=" + row.getColLast());
                for (ExcelCell cell : cells) {
                    System.out.println("col： " + cell.getColIndex() + " ，value = " + cell.getStrValue());
                }
            }

            @Override
            public void onSheetEnd(@Nonnull ExcelSheet sheet) {
                System.out.println("结束sheet：" + sheet.getShtName());
            }
        });
    }
}