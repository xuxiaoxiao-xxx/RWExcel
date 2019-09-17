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
        Workbook workbook = provider.version() == Type.XLS ? new HSSFWorkbook() : new SXSSFWorkbook(100);

        int lastShtIndex = -1;
        for (int i = 0; ; i++) {
            ExcelSheet excelSheet = provider.nextSheet(lastShtIndex);
            if (excelSheet == null) {
                break;
            } else if (excelSheet.getShtIndex() != i) {
                throw new IllegalArgumentException(String.format("写出sheet错误，期望shtIndex：%d，实际shtIndex：%d", i, excelSheet.getShtIndex()));
            } else if (excelSheet.getShtName().trim().isEmpty()) {
                throw new IllegalArgumentException("写出sheet错误，shtName不能为空");
            } else {
                lastShtIndex = i;
                Sheet sheet = workbook.createSheet(excelSheet.getShtName());
                workbook.setSheetOrder(excelSheet.getShtName(), excelSheet.getShtIndex());

                int lastRowIndex = -1;
                for (int j = 0; ; j++) {
                    ExcelRow excelRow = provider.nextRow(i, lastRowIndex);
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
                        for (int k = excelRow.getColFirst(); k <= excelRow.getColLast(); k++) {
                            ExcelCell excelCell = provider.provideCell(i, j, k);
                            if (excelCell != null) {
                                if (excelCell.getShtIndex() != i) {
                                    throw new IllegalArgumentException(String.format("写出cell错误，期望shtIndex：%d，实际shtIndex：%d", i, excelCell.getShtIndex()));
                                } else if (excelCell.getRowIndex() != j) {
                                    throw new IllegalArgumentException(String.format("写出cell错误，期望rowIndex：%d，实际rowIndex：%d", j, excelCell.getRowIndex()));
                                } else if (excelCell.getColIndex() != k) {
                                    throw new IllegalArgumentException(String.format("写出cell错误，期望colIndex：%d，实际colIndex：%d", j, excelCell.getColIndex()));
                                } else {
                                    Cell cell = row.createCell(k, CellType.STRING);
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
    }
}
