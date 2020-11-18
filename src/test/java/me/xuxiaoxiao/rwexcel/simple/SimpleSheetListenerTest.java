package me.xuxiaoxiao.rwexcel.simple;

import me.xuxiaoxiao.rwexcel.reader.ExcelReader;
import me.xuxiaoxiao.rwexcel.reader.ExcelStreamReader;
import me.xuxiaoxiao.rwexcel.reader.ExcelUserReader;
import org.junit.Test;

import java.util.Objects;

public class SimpleSheetListenerTest {
    @Test
    public void testSimpleSheetListener() throws Exception {
        TestSimpleListener listener = new TestSimpleListener();

        ExcelReader streamReader = new ExcelStreamReader();
        streamReader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("testSimpleSheetListener.xls")), listener);
        streamReader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("testSimpleSheetListener.xlsx")), listener);

        ExcelReader userReader = new ExcelUserReader();
        userReader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("testSimpleSheetListener.xls")), listener);
        userReader.read(Objects.requireNonNull(Thread.currentThread().getContextClassLoader().getResourceAsStream("testSimpleSheetListener.xlsx")), listener);
    }
}