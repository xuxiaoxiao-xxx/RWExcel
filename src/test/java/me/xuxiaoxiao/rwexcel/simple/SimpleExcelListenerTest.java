package me.xuxiaoxiao.rwexcel.simple;

import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.reader.ExcelReader;
import me.xuxiaoxiao.rwexcel.reader.ExcelStreamReader;
import me.xuxiaoxiao.rwexcel.reader.ExcelUserReader;
import me.xuxiaoxiao.rwexcel.reader.TestListener;
import org.junit.Test;

import java.util.Objects;

public class SimpleExcelListenerTest {
    @Test
    public void testSimpleExcelListener() throws Exception {
        final TestSimpleListener testSimpleListener = new TestSimpleListener();
        final TestListener testListener = new TestListener();
        final ExcelReader.Listener listener = new SimpleExcelListener() {
            @Override
            public ExcelReader.Listener sheetListener(ExcelSheet sheet) {
                if (sheet.getShtIndex() == 0) {
                    return testSimpleListener;
                } else if (sheet.getShtIndex() == 1) {
                    return testListener;
                } else {
                    return null;
                }
            }
        };

        ExcelReader streamReader = new ExcelStreamReader();
        streamReader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("testSimpleExcelListener.xls")), listener);
        streamReader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("testSimpleExcelListener.xlsx")), listener);
        ExcelReader userReader = new ExcelUserReader();

        userReader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("testSimpleExcelListener.xls")), listener);
        userReader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("testSimpleExcelListener.xlsx")), listener);
    }
}