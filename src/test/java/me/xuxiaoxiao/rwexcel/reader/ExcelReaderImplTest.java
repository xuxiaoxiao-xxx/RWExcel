package me.xuxiaoxiao.rwexcel.reader;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import org.junit.Test;

import javax.annotation.Nonnull;
import java.util.Objects;

public class ExcelReaderImplTest {

    @Test
    public void read() throws Exception {
        ExcelReader reader = new ExcelReaderImpl();
        reader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("test.xlsx")), new ExcelReader.Listener() {
            private String sheetName;

            @Override
            public void onSheet(@Nonnull ExcelSheet sheet) {
                System.out.println();
                System.out.println("获取到sheet：" + sheet.getShtName());
                sheetName = sheet.getShtName();
            }

            @Override
            public void onRow(@Nonnull ExcelRow row) {
                System.out.println();
                System.out.println("获取到row：sheet=" + sheetName + "，row=" + row.getRowIndex() + " ，firstCol=" + row.getColFirst() + " ，lastCol=" + row.getColLast());

            }

            @Override
            public void onCell(@Nonnull ExcelCell cell) {
                System.out.println("获取到col：sheet=" + sheetName + "，row=" + cell.getRowIndex() + " ，col = " + cell.getColIndex() + " ，value = " + cell.getStrValue());
            }
        });
    }
}