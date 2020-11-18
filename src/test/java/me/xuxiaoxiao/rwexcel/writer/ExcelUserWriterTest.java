package me.xuxiaoxiao.rwexcel.writer;

import me.xuxiaoxiao.rwexcel.reader.ExcelReader;
import me.xuxiaoxiao.rwexcel.reader.ExcelStreamReader;
import me.xuxiaoxiao.rwexcel.reader.TestListener;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;

public class ExcelUserWriterTest {

    @Test
    public void writeXls() throws Exception {
        ByteArrayOutputStream baOutStream = new ByteArrayOutputStream();

        ExcelWriter excelWriter = new ExcelUserWriter();
        excelWriter.write(baOutStream, new TestProvider(ExcelWriter.Version.XLS));

        ExcelReader excelReader = new ExcelStreamReader();
        excelReader.read(new ByteArrayInputStream(baOutStream.toByteArray()), new TestListener());
    }

    @Test
    public void writeXlsx() throws Exception {
        ByteArrayOutputStream baOutStream = new ByteArrayOutputStream();

        ExcelWriter excelWriter = new ExcelUserWriter();
        excelWriter.write(baOutStream, new TestProvider(ExcelWriter.Version.XLSX));

        ExcelReader excelReader = new ExcelStreamReader();
        excelReader.read(new ByteArrayInputStream(baOutStream.toByteArray()), new TestListener());
    }
}