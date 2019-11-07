package me.xuxiaoxiao.rwexcel.reader;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import org.apache.poi.ss.usermodel.*;

import javax.annotation.Nonnull;
import java.io.InputStream;
import java.util.LinkedList;
import java.util.List;

/**
 * Excel解析器UserModel实现
 * <ul>
 * <li>[2019/11/7 12:44]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public class ExcelUserReader implements ExcelReader {

    @Override
    public void read(@Nonnull InputStream inStream, @Nonnull Listener listener) throws Exception {
        Workbook workbook = WorkbookFactory.create(inStream);
        for (int i = 0, count = workbook.getNumberOfSheets(); i < count; i++) {
            Sheet sheet = workbook.getSheetAt(i);
            ExcelSheet excelSheet = new ExcelSheet(i, sheet.getSheetName());

            listener.onSheetStart(excelSheet);

            DataFormatter formatter = new DataFormatter();
            for (int j = sheet.getFirstRowNum(); j < sheet.getLastRowNum() + 1; j++) {
                Row row = sheet.getRow(j);
                if (row != null) {
                    List<ExcelCell> excelCells = new LinkedList<>();
                    for (int k = row.getFirstCellNum(); k < row.getLastCellNum() + 1; k++) {
                        Cell cell = row.getCell(k);

                        String value = formatter.formatCellValue(cell);
                        if (value != null && value.length() > 0) {
                            excelCells.add(new ExcelCell(i, j, k, value));
                        }
                    }

                    if (!excelCells.isEmpty()) {
                        ExcelRow excelRow = new ExcelRow(i, j);
                        excelRow.setColFirst(excelCells.get(0).getColIndex());
                        excelRow.setColLast(excelCells.get(excelCells.size() - 1).getColIndex());
                        listener.onRow(excelSheet, excelRow, excelCells);
                    }
                }
            }

            listener.onSheetEnd(excelSheet);
        }
    }
}
