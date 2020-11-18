package me.xuxiaoxiao.rwexcel.reader;

import org.junit.Test;

import java.util.Objects;

public class ExcelStreamReaderTest {

    @Test
    public void readXls() throws Exception {
        ExcelReader reader = new ExcelStreamReader();
        TestListener listener = new TestListener();
        reader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("testListener.xls")), listener);
    }

    @Test
    public void readXlsx() throws Exception {
        ExcelReader reader = new ExcelStreamReader();
        TestListener listener = new TestListener();
        reader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("testListener.xlsx")), listener);
    }
}