package me.xuxiaoxiao.rwexcel.simple;

import lombok.Getter;
import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.reader.ExcelReader;
import me.xuxiaoxiao.rwexcel.simple.converter.Converter;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.List;
import java.util.TreeMap;

/**
 * 简单Sheet监听器，将sheet数据解析成List
 * <ul>
 * <li>[2019/9/19 16:41]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public abstract class SimpleSheetListener<T> implements ExcelReader.Listener {
    private final List<T> list;

    @Getter
    private final ExcelSheet sheet;
    @Getter
    private final int cache;
    @Getter
    private final Class<T> clazz;
    @Getter
    private final boolean adaptive;

    private int rowStart = 0, rowEnd = 0;

    protected final TreeMap<Integer, Field> fMapper = new TreeMap<>();
    protected final TreeMap<Integer, Converter> cMapper = new TreeMap<>();

    /**
     * 创建一个sheet监听器，指定默认100的列表缓存大小
     *
     * @param sheet 当前sheet信息
     */
    public SimpleSheetListener(ExcelSheet sheet) {
        this(sheet, 100);
    }


    /**
     * 创建一个sheet监听器，指定列表缓存大小
     *
     * @param sheet 当前sheet信息
     */
    public SimpleSheetListener(ExcelSheet sheet, int cache) {
        this.list = new ArrayList<>(cache);
        this.sheet = sheet;
        this.cache = cache;
        this.clazz = entityClass();

        boolean adaptive = true;
        for (Class<?> i = this.clazz; !i.equals(Object.class); i = i.getSuperclass()) {
            Field[] fields = i.getDeclaredFields();
            for (Field field : fields) {
                if (field.isAnnotationPresent(ExcelColumn.class)) {
                    field.setAccessible(true);

                    ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                    if (excelColumn.index() >= 0) {
                        adaptive = false;
                        if (fMapper.containsKey(excelColumn.index())) {
                            throw new IllegalArgumentException(String.format("ExcelColumn列冲突，有多个序号为 %d 的列", excelColumn.index()));
                        } else {
                            fMapper.put(excelColumn.index(), field);
                        }
                    } else if (!adaptive) {
                        throw new IllegalArgumentException(String.format("ExcelColumn的列 %s 未设置序号", excelColumn.value()));
                    }
                }
            }
        }
        this.adaptive = adaptive;
    }

    @Override
    public final void onSheetStart(@Nonnull ExcelSheet sheet) {
    }

    @Override
    public final void onRow(@Nonnull ExcelSheet sheet, @Nonnull ExcelRow row, @Nonnull List<ExcelCell> cells) {
        if (row.getRowIndex() < titleRowCount()) {
            titleRow(row.getRowIndex(), row, cells);
            rowStart = row.getRowIndex() + 1;
            rowEnd = row.getRowIndex() + 1;
        } else {
            while (rowEnd < row.getRowIndex()) {
                try {
                    //处理空白行
                    if (!rowSkip(rowEnd, null, null)) {
                        list.add(rowEntity(rowEnd, null, null));
                        if (list.size() >= cache) {
                            flush();
                        }
                    }
                    rowEnd++;
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
            }
            try {
                if (!rowSkip(row.getRowIndex(), row, cells)) {
                    list.add(rowEntity(row.getRowIndex(), row, cells));
                    if (list.size() >= cache) {
                        flush();
                    }
                }
                rowEnd++;
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        }
    }

    @Override
    public final void onSheetEnd(@Nonnull ExcelSheet sheet) {
        flush();
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
     * 标题行解析
     *
     * @param rowIndex 行号
     * @param row      标题行信息
     * @param cells    标题行单元格信息
     */
    protected void titleRow(int rowIndex, @Nullable ExcelRow row, @Nullable List<ExcelCell> cells) {
        if (rowIndex == 0 && adaptive && row != null && cells != null && !cells.isEmpty()) {
            //如果是自动寻列模式，要找一下属性和列的对应关系
            for (Class<?> i = this.clazz; !i.equals(Object.class); i = i.getSuperclass()) {
                Field[] fields = i.getDeclaredFields();
                for (Field field : fields) {
                    if (field.isAnnotationPresent(ExcelColumn.class)) {
                        ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                        for (ExcelCell cell : cells) {
                            if (excelColumn.value().equals(cell.getStrValue())) {
                                if (fMapper.containsKey(cell.getColIndex())) {
                                    throw new IllegalArgumentException(String.format("ExcelColumn自动寻列冲突，有多个成员域title为 %s", excelColumn.value()));
                                } else {
                                    fMapper.put(cell.getColIndex(), field);
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    private void flush() {
        onList(rowStart, rowEnd, list);
        rowStart = rowEnd;
        list.clear();
    }

    /**
     * 获取实体类的class对象
     *
     * @return 实体类的class对象
     */
    @Nonnull
    protected Class<T> entityClass() {
        if (this.clazz == null) {
            Class<?> listenerClass = this.getClass();
            if (listenerClass.getGenericSuperclass() instanceof ParameterizedType) {
                Type[] typeArguments = ((ParameterizedType) listenerClass.getGenericSuperclass()).getActualTypeArguments();
                if (typeArguments != null && typeArguments.length > 0) {
                    Type paramType = typeArguments[0];
                    if (paramType instanceof Class) {
                        return (Class<T>) paramType;
                    }
                }
            }
            throw new IllegalStateException("未能自动识别SimpleExcelListener的模板参数");
        } else {
            return this.clazz;
        }
    }

    /**
     * 获取excel列对应的Field对象
     *
     * @param rowIndex 行号
     * @param colIndex 列号
     * @return 对应的Field对象
     */
    @Nullable
    protected Field entityField(int rowIndex, int colIndex) {
        return fMapper.get(colIndex);
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
     * 是否跳过某行
     *
     * @param rowIndex 行号
     * @param row      行信息
     * @param cells    行内所有单元格
     * @return 是否跳过
     * @throws Exception 判断过程中出现的异常
     */
    protected boolean rowSkip(int rowIndex, @Nullable ExcelRow row, @Nullable List<ExcelCell> cells) throws Exception {
        //第一行和空白行都要跳过
        return rowIndex < titleRowCount() || row == null;
    }

    /**
     * 获取某行对应的实体信息
     *
     * @param rowIndex 行号
     * @param row      行信息
     * @param cells    行内所有单元格
     * @return 对应的实体信息
     * @throws Exception 解析过程中出现的异常
     */
    @Nullable
    protected T rowEntity(int rowIndex, @Nullable ExcelRow row, @Nullable List<ExcelCell> cells) throws Exception {
        if (row == null || cells == null || cells.isEmpty()) {
            return null;
        } else {
            T entity = this.clazz.newInstance();
            for (ExcelCell cell : cells) {
                Field field = entityField(cell.getRowIndex(), cell.getColIndex());
                if (field != null) {
                    Converter converter = cMapper.get(cell.getColIndex());
                    if (converter == null) {
                        converter = field.getAnnotation(ExcelColumn.class).converter().newInstance();
                        cMapper.put(cell.getColIndex(), converter);
                    }
                    try {
                        field.setAccessible(true);
                        field.set(entity, converter.str2obj(field, cell.getStrValue()));
                    } catch (Exception e) {
                        if (mismatchPolicy() == Converter.MismatchPolicy.Throw) {
                            throw e;
                        }
                    }
                }
            }
            return entity;
        }
    }

    /**
     * 解析到部分数据，进行一个批次的处理，会调用多次
     *
     * @param rowStart 解析的开始行数，包含
     * @param rowEnd   解析的结束行数，不包含
     * @param list     解析到的列表
     */
    protected abstract void onList(int rowStart, int rowEnd, @Nonnull List<T> list);
}
