package me.xuxiaoxiao.rwexcel.simple;

import lombok.Getter;
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
 * 简单Sheet数据源，将List转换成sheet数据
 * <ul>
 * <li>[2019/9/20 13:33]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public abstract class SimpleSheetProvider<T> implements ExcelWriter.Provider {
    /**
     * 模板类型
     */
    @Getter
    private final Class<T> clazz;
    /**
     * Java类属性和Excel列号的映射关系
     */
    @Getter
    private final Map<Field, Integer> mapper = new HashMap<>();
    /**
     * Java类属性和属性转换器的映射关系
     */
    @Getter
    private final Map<Field, Converter> converters = new HashMap<>();
    /**
     * Excel第一列的列号
     */
    @Getter
    private final int colFirst;
    /**
     * Excel最后一列的列号
     */
    @Getter
    private final int colLast;

    private final List<T> list = new LinkedList<>();

    /**
     * 创建一个sheet数据源
     */
    public SimpleSheetProvider() {
        this.clazz = detectEntityClass();
        this.mapper.putAll(detectEntityMapper());

        int colF = Integer.MAX_VALUE, colL = Integer.MIN_VALUE;
        for (Map.Entry<Field, Integer> entry : mapper.entrySet()) {
            if (colF > entry.getValue()) {
                colF = entry.getValue();
            }
            if (colL < entry.getValue()) {
                colL = entry.getValue();
            }
        }
        this.colFirst = colF;
        this.colLast = colL;
    }

    /**
     * 获取模板类型的class对象
     *
     * @return 模板类型的class对象
     */
    @Nonnull
    @SuppressWarnings("unchecked")
    protected Class<T> detectEntityClass() {
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
        throw new IllegalStateException("未能自动识别SimpleExcelListener的模板参数，请勿多级继承SimpleExcelListener，或自行实现detectEntityClass方法");
    }

    /**
     * 获取Excel列号和Java类属性的映射关系
     *
     * @return Excel列号和Java类属性的映射关系
     */
    @Nonnull
    protected Map<Field, Integer> detectEntityMapper() {
        Map<Field, Integer> map = new HashMap<>();

        for (Class<?> i = this.clazz; !i.equals(Object.class); i = i.getSuperclass()) {
            Field[] fields = i.getDeclaredFields();
            for (Field field : fields) {
                if (field.isAnnotationPresent(ExcelColumn.class)) {
                    field.setAccessible(true);

                    ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                    if (excelColumn.index() >= 0) {
                        //模板类型class设置了index，将映射关系存入到mapper中
                        for (Map.Entry<Field, Integer> entry : map.entrySet()) {
                            if (entry.getValue().equals(excelColumn.index())) {
                                throw new IllegalArgumentException(String.format("%s 中有多个属性映射到了Excel的第 %d 列", this.clazz.getSimpleName(), excelColumn.index()));
                            }
                        }
                        map.put(field, excelColumn.index());
                    } else {
                        //模板类型class没有设置index，根据列名将映射关系存入到mapper中
                        int index = 0;
                        while (true) {
                            boolean exist = false;
                            for (Map.Entry<Field, Integer> entry : map.entrySet()) {
                                if (entry.getValue().equals(index)) {
                                    exist = true;
                                    break;
                                }
                            }
                            if (exist) {
                                index++;
                            } else {
                                map.put(field, index);
                                break;
                            }
                        }
                    }
                }
            }
        }
        return map;
    }

    @Nonnull
    @Override
    public ExcelWriter.Version version() {
        return ExcelWriter.Version.XLSX;
    }

    @Nullable
    @Override
    public ExcelSheet provideSheet(int lastSheetIndex) {
        if (lastSheetIndex < 0) {
            return new ExcelSheet(0, "Sheet1");
        } else {
            return null;
        }
    }

    @Nullable
    @Override
    public final ExcelRow provideRow(@Nonnull ExcelSheet sheet, int lastRowIndex) {
        if (lastRowIndex + 1 < titleRowCount()) {
            return titleRow(sheet, lastRowIndex);
        } else {
            if (list.isEmpty()) {
                List<T> temp = queryList(lastRowIndex);
                if (temp != null) {
                    this.list.addAll(temp);
                }
            }
            while (list.size() > 0 && entitySkip(sheet, lastRowIndex, list.get(0))) {
                list.remove(0);
            }
            if (list.isEmpty()) {
                return null;
            } else {
                return entityRow(sheet, lastRowIndex, list.get(0));
            }
        }
    }

    @Nonnull
    @Override
    public final List<ExcelCell> provideCells(@Nonnull ExcelSheet sheet, @Nonnull ExcelRow row) {
        if (row.getRowIndex() < titleRowCount()) {
            return titleCells(sheet, row);
        } else {
            try {
                return entityCells(sheet, row, list.remove(0));
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
    }

    /**
     * 转换异常的处理方法
     *
     * @return 转换异常的处理方法
     */
    @Nonnull
    public Converter.MismatchPolicy mismatchPolicy() {
        return Converter.MismatchPolicy.Throw;
    }

    /**
     * 标题行数量
     *
     * @return 标题行数量
     */
    protected int titleRowCount() {
        return 1;
    }

    /**
     * 标题行信息
     *
     * @param sheet        当前sheet
     * @param lastRowIndex 上一个行号
     * @return 标题行信息
     */
    @Nullable
    protected ExcelRow titleRow(@Nonnull ExcelSheet sheet, int lastRowIndex) {
        ExcelRow excelRow = new ExcelRow(sheet.getShtIndex(), lastRowIndex + 1);
        excelRow.setColFirst(this.colFirst);
        excelRow.setColLast(this.colLast);
        return excelRow;
    }

    /**
     * 标题行的单元格列表
     *
     * @param sheet 当前sheet
     * @param row   标题行信息
     * @return 标题行的单元格列表
     */
    @Nonnull
    protected List<ExcelCell> titleCells(@Nonnull ExcelSheet sheet, @Nonnull ExcelRow row) {
        List<ExcelCell> cells = new LinkedList<>();
        for (Map.Entry<Field, Integer> entry : this.mapper.entrySet()) {
            cells.add(new ExcelCell(sheet.getShtIndex(), row.getRowIndex(), entry.getValue(), entry.getKey().getAnnotation(ExcelColumn.class).value()));
        }
        Collections.sort(cells, new Comparator<ExcelCell>() {
            @Override
            public int compare(ExcelCell o1, ExcelCell o2) {
                return o1.getColIndex() - o2.getColIndex();
            }
        });
        return cells;
    }

    /**
     * 是否要跳过某个实体
     *
     * @param sheet        当前sheet
     * @param lastRowIndex 上一个行号
     * @param entity       实体信息
     * @return 是否要跳过
     */
    protected boolean entitySkip(@Nonnull ExcelSheet sheet, int lastRowIndex, @Nullable T entity) {
        //跳过null实体
        return entity == null;
    }

    /**
     * 获取实体对应的行信息
     *
     * @param sheet        当前sheet
     * @param lastRowIndex 上一个行号
     * @param entity       实体信息
     * @return 实体对应的行信息
     */
    @Nullable
    protected ExcelRow entityRow(@Nonnull ExcelSheet sheet, int lastRowIndex, @Nullable T entity) {
        ExcelRow excelRow = new ExcelRow(sheet.getShtIndex(), lastRowIndex + 1);
        excelRow.setColFirst(this.colFirst);
        excelRow.setColLast(this.colLast);
        return excelRow;
    }

    /**
     * 获取实体对应的单元格列表
     *
     * @param sheet  当前sheet
     * @param row    实体对应的行信息
     * @param entity 实体信息
     * @return 实体对应的单元格列表
     * @throws Exception 转换时可能发生异常
     */
    @Nonnull
    protected List<ExcelCell> entityCells(@Nonnull ExcelSheet sheet, @Nonnull ExcelRow row, @Nullable T entity) throws Exception {
        if (entity == null) {
            return new LinkedList<>();
        } else {
            List<ExcelCell> cells = new LinkedList<>();
            for (Map.Entry<Field, Integer> entry : this.mapper.entrySet()) {
                Field field = entry.getKey();
                Converter converter = converters.get(field);
                if (converter == null) {
                    converter = field.getAnnotation(ExcelColumn.class).converter().newInstance();
                    converters.put(field, converter);
                }
                try {
                    cells.add(new ExcelCell(sheet.getShtIndex(), row.getRowIndex(), mapper.get(field), converter.obj2str(field, field.get(entity))));
                } catch (Exception e) {
                    if (mismatchPolicy() == Converter.MismatchPolicy.Throw) {
                        throw e;
                    }
                }
            }
            Collections.sort(cells, new Comparator<ExcelCell>() {
                @Override
                public int compare(ExcelCell o1, ExcelCell o2) {
                    return o1.getColIndex() - o2.getColIndex();
                }
            });
            return cells;
        }
    }

    /**
     * 查询实体列表，进行一个批次的处理，会调用多次
     *
     * @param lastRowIndex 最后一行的行号，首次为-1
     * @return 一批次的实体列表
     */
    public abstract List<T> queryList(int lastRowIndex);
}
