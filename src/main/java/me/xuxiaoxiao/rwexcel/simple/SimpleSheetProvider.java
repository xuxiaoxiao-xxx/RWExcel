package me.xuxiaoxiao.rwexcel.simple;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.simple.converter.Converter;
import me.xuxiaoxiao.rwexcel.writer.ExcelWriter;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.*;

/**
 * 请填写类的描述
 * <ul>
 * <li>[2019/9/20 13:33]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public abstract class SimpleSheetProvider<T> implements ExcelWriter.Provider {
    private final ExcelSheet sheet;
    protected final int colFirst, colLast;
    private final List<T> list = new LinkedList<>();
    protected final Converter converter = new Converter();
    protected final Map<Field, Integer> fMapper = new HashMap<>();
    protected final Map<Field, Converter> cMapper = new HashMap<>();
    private final Class<T> clazz;


    public SimpleSheetProvider(ExcelSheet sheet) {
        this.sheet = sheet;
        this.clazz = entityClass();

        int colF = Integer.MAX_VALUE, colL = Integer.MIN_VALUE;
        boolean adaptive = true;
        List<Field> fieldList = new LinkedList<>();
        for (Class<?> i = this.clazz; !i.equals(Object.class); i = i.getSuperclass()) {
            Field[] fields = i.getDeclaredFields();
            for (Field field : fields) {
                if (field.isAnnotationPresent(ExcelColumn.class)) {
                    field.setAccessible(true);
                    fieldList.add(field);

                    ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                    if (excelColumn.index() >= 0) {
                        adaptive = false;
                        if (excelColumn.index() < colF) {
                            colF = excelColumn.index();
                        }
                        if (colL < excelColumn.index()) {
                            colL = excelColumn.index();
                        }
                        fMapper.put(field, excelColumn.index());
                    } else if (!adaptive) {
                        throw new IllegalArgumentException(String.format("ExcelColumn的列 %s 未设置序号", excelColumn.value()));
                    }
                }
            }
        }
        if (adaptive) {
            Collections.sort(fieldList, new Comparator<Field>() {
                @Override
                public int compare(Field o1, Field o2) {
                    Class<?> c1 = o1.getDeclaringClass();
                    Class<?> c2 = o2.getDeclaringClass();
                    if (c1.equals(c2)) {
                        return 0;
                    } else if (c1.isAssignableFrom(c2)) {
                        return -1;
                    } else {
                        return 1;
                    }
                }
            });
            colF = 0;
            colL = fieldList.size() - 1;
            for (int i = 0; i < fieldList.size(); i++) {
                fMapper.put(fieldList.get(i), i);
            }
        }
        this.colFirst = colF;
        this.colLast = colL;
        if (fMapper.isEmpty()) {
            throw new IllegalArgumentException("SimpleSheetProvider的模板参数中未找到任何使用@Excelolumn标记的属性");
        }
    }

    @Nonnull
    @Override
    public ExcelWriter.Type version() {
        return ExcelWriter.Type.XLSX;
    }

    @Nonnull
    @Override
    public ExcelSheet[] sheets() {
        return new ExcelSheet[0];
    }

    @Nullable
    @Override
    public final ExcelRow provideRow(ExcelSheet sheet, int lastRowIndex) {
        if (list.isEmpty()) {
            List<T> temp = queryList(this.sheet, lastRowIndex);
            if (temp != null) {
                this.list.addAll(temp);
            }
        }
        if (list.isEmpty()) {
            return null;
        } else {
            return entityRow(sheet, lastRowIndex, list.get(0));
        }
    }

    @Nonnull
    @Override
    public final List<ExcelCell> provideCells(ExcelSheet sheet, ExcelRow row) {
        return entityCells(sheet, row, list.remove(0));
    }

    protected Class<T> entityClass() {
        if (this.clazz == null) {
            Class<?> providerClass = this.getClass();
            if (providerClass.getGenericSuperclass() instanceof ParameterizedType) {
                Type[] typeArguments = ((ParameterizedType) providerClass.getGenericSuperclass()).getActualTypeArguments();
                if (typeArguments != null && typeArguments.length > 0) {
                    Type paramType = typeArguments[0];
                    if (paramType instanceof Class) {
                        return (Class<T>) paramType;
                    }
                }
            }
            throw new IllegalStateException("未能自动识别SimpleSheetProvider的模板参数");
        } else {
            return this.clazz;
        }
    }

    protected ExcelRow entityRow(ExcelSheet sheet, int lastRowIndex, T entity) {
        ExcelRow excelRow = new ExcelRow(sheet.getShtIndex(), lastRowIndex + 1);
        excelRow.setColFirst(this.colFirst);
        excelRow.setColLast(this.colLast);
        return excelRow;
    }

    protected List<ExcelCell> entityCells(ExcelSheet sheet, ExcelRow row, T entity) {
        List<ExcelCell> cells = new LinkedList<>();
        for (Map.Entry<Field, Integer> entry : this.fMapper.entrySet()) {
            try {
                Field field = entry.getKey();
                Class<? extends Converter> cClass = field.getAnnotation(ExcelColumn.class).converter();
                if (cClass.equals(Converter.class)) {
                    cells.add(new ExcelCell(sheet.getShtIndex(), row.getRowIndex(), entry.getValue(), converter.obj2str(field, field.get(entity))));
                } else {
                    Converter converter = cMapper.get(field);
                    if (converter == null) {
                        converter = cClass.newInstance();
                        cMapper.put(field, converter);
                    }
                    cells.add(new ExcelCell(sheet.getShtIndex(), row.getRowIndex(), entry.getValue(), converter.obj2str(field, field.get(entity))));
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return cells;
    }

    public abstract List<T> queryList(ExcelSheet sheet, int lastRowIndex);
}
