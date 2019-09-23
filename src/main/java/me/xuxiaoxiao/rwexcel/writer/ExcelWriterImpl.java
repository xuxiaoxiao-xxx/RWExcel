package me.xuxiaoxiao.rwexcel.writer;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.annotation.Nonnull;
import java.io.OutputStream;

/**
 * Excel导出器实现，03版本Excel不支持流式写出，07版本支持流式写出
 * <ul>
 * <li>[2019/9/12 19:42]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public class ExcelWriterImpl implements ExcelWriter {

    @Override
    public void write(@Nonnull OutputStream outStream, @Nonnull Provider provider) throws Exception {
        Workbook workbook = provider.version() == Version.XLS ? new HSSFWorkbook() : new SXSSFWorkbook(100);

        for (int i = 0; ; i++) {
            ExcelSheet excelSheet = provider.provideSheet(i - 1);
            if (excelSheet == null) {
                break;
            } else if (excelSheet.getShtIndex() != i) {
                throw new IllegalArgumentException(String.format("写出sheet错误，期望shtIndex：%d，实际shtIndex：%d", i, excelSheet.getShtIndex()));
            } else if (excelSheet.getShtName().trim().isEmpty()) {
                throw new IllegalArgumentException("写出sheet错误，shtName不能为空");
            } else {
                Sheet sheet = workbook.createSheet(excelSheet.getShtName());
                workbook.setSheetOrder(excelSheet.getShtName(), excelSheet.getShtIndex());

                int lastRowIndex = -1;
                for (int j = 0; ; j++) {
                    ExcelRow excelRow = provider.provideRow(excelSheet, lastRowIndex);
                    if (excelRow == null) {
                        break;
                    } else if (excelRow.getShtIndex() != i) {
                        throw new IllegalArgumentException(String.format("写出row错误，期望shtIndex：%d，实际shtIndex：%d", i, excelRow.getShtIndex()));
                    } else if (excelRow.getRowIndex() < j) {
                        throw new IllegalArgumentException(String.format("写出row错误，期望rowIndex：至少为%d，实际shtIndex：%d", j, excelRow.getShtIndex()));
                    } else {
                        while (j < excelRow.getRowIndex()) {
                            sheet.createRow(j++);
                        }
                        lastRowIndex = j;
                        Row row = sheet.createRow(excelRow.getRowIndex());

                        int lastColIndex = -1;
                        for (ExcelCell excelCell : provider.provideCells(excelSheet, excelRow)) {
                            if (excelCell != null) {
                                if (excelCell.getShtIndex() != i) {
                                    throw new IllegalArgumentException(String.format("写出cell错误，期望shtIndex：%d，实际shtIndex：%d", i, excelCell.getShtIndex()));
                                } else if (excelCell.getRowIndex() != j) {
                                    throw new IllegalArgumentException(String.format("写出cell错误，期望rowIndex：%d，实际rowIndex：%d", j, excelCell.getRowIndex()));
                                } else if (excelCell.getColIndex() < 0 || excelCell.getColIndex() <= lastColIndex) {
                                    throw new IllegalArgumentException(String.format("写出cell错误，期望colIndex大于0且按顺序，实际colIndex：%d", excelCell.getColIndex()));
                                } else {
                                    lastColIndex = excelCell.getColIndex();
                                    Cell cell = row.createCell(excelCell.getColIndex(), CellType.STRING);
                                    cell.setCellValue(excelCell.getStrValue());
                                }
                            }
                        }
                    }
                }
            }
        }
        workbook.write(outStream);
        if (workbook instanceof SXSSFWorkbook) {
            ((SXSSFWorkbook) workbook).dispose();
        }
        workbook.close();
        outStream.close();
    }
}
