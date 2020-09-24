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
import java.util.*;

/**
 * 简单Sheet监听器，将sheet数据解析成List
 * <ul>
 * <li>[2019/9/19 16:41]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public abstract class SimpleSheetListener<T> implements ExcelReader.Listener {
    /**
     * 模板类型
     */
    @Getter
    private final Class<T> clazz;
    /**
     * 缓存大小
     */
    @Getter
    private final int cache;

    /**
     * Excel列号和Java类属性的映射关系
     */
    @Getter
    private final TreeMap<Integer, Field> mapper = new TreeMap<>();
    /**
     * Excel列号和属性转换器的映射关系
     */
    @Getter
    private final TreeMap<Integer, Converter> converters = new TreeMap<>();
    /**
     * 所有的标题行
     */
    private final List<List<ExcelCell>> titles = new ArrayList<>();
    /**
     * 当前批次缓存数据
     */
    private final List<T> list = new ArrayList<>();
    /**
     * 当前处理的开始行（包含），和结束行（不包含）
     */
    private int rowStart = 0, rowEnd = 0;

    /**
     * 创建一个sheet监听器，指定列表缓存大小为100
     */
    public SimpleSheetListener() {
        this(100);
    }

    /**
     * 创建一个sheet监听器，指定列表缓存大小
     *
     * @param cache 缓存大小
     */
    public SimpleSheetListener(int cache) {
        this.cache = cache;
        this.clazz = detectEntityClass();
    }

    /**
     * 重置监听器，以便监听器重复使用
     */
    private void reset() {
        titles.clear();
        mapper.clear();
        list.clear();
        rowStart = 0;
        rowEnd = 0;
    }

    /**
     * 处理并清空缓存
     */
    private void flush() {
        if (!list.isEmpty()) {
            onList(rowStart, rowEnd, list);
            list.clear();
        }
        rowStart = rowEnd;
    }

    /**
     * 获取模板类型的class对象
     *
     * @return 模板类型的class对象
     */
    @Nonnull
    @SuppressWarnings("unchecked")
    protected Class<T> detectEntityClass() {
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
        throw new IllegalStateException("未能自动识别SimpleExcelListener的模板参数，请勿多级继承SimpleExcelListener，或自行实现detectEntityClass方法");
    }

    /**
     * 获取Excel列号和Java类属性的映射关系
     *
     * @param titles 所有标题行
     * @return Excel列号和Java类属性的映射关系
     */
    @Nonnull
    protected Map<Integer, Field> detectEntityMapper(List<List<ExcelCell>> titles) {
        Map<Integer, Field> map = new HashMap<>();
        List<ExcelCell> titleRow = titles.size() > 0 ? titles.get(titles.size() - 1) : new LinkedList<ExcelCell>();

        for (Class<?> i = this.clazz; !i.equals(Object.class); i = i.getSuperclass()) {
            Field[] fields = i.getDeclaredFields();
            for (Field field : fields) {
                if (field.isAnnotationPresent(ExcelColumn.class)) {
                    field.setAccessible(true);

                    int index = -1;
                    ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                    if (excelColumn.index() >= 0) {
                        //模板类型class设置了index，将映射关系存入到mapper中
                        index = excelColumn.index();
                    } else {
                        //模板类型class没有设置index，根据列名将映射关系存入到mapper中
                        for (ExcelCell titleCell : titleRow) {
                            if (excelColumn.value().equals(titleCell.getStrValue())) {
                                index = titleCell.getColIndex();
                                break;
                            }
                        }
                    }

                    if (index >= 0) {
                        if (map.containsKey(index)) {
                            throw new IllegalArgumentException(String.format("%s 中有多个属性映射到了Excel的第 %d 列", this.clazz.getSimpleName(), index));
                        } else {
                            map.put(index, field);
                        }
                    }
                }
            }
        }
        return map;
    }

    @Override
    public final void onSheetStart(@Nonnull ExcelSheet sheet) {
        reset();
    }

    @Override
    public final void onRow(@Nonnull ExcelSheet sheet, @Nonnull ExcelRow row, @Nonnull List<ExcelCell> cells) {
        if (handleSheet(sheet)) {
            if (row.getRowIndex() < titleRowCount()) {
                //当前行是标题行
                titles.add(cells);
                rowStart = row.getRowIndex() + 1;
                rowEnd = row.getRowIndex() + 1;
            } else {
                if (mapper.isEmpty()) {
                    mapper.putAll(detectEntityMapper(titles));
                }
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
    }

    @Override
    public final void onSheetEnd(@Nonnull ExcelSheet sheet) {
        flush();
    }

    /**
     * 是否要处理某个sheet
     *
     * @param sheet sheet
     * @return 是否要处理
     */
    protected boolean handleSheet(@Nonnull ExcelSheet sheet) {
        return true;
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
     * 列转换属性时 发生异常的处理策略
     *
     * @return 处理策略
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
        //标题行和空白行都要跳过
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
                Field field = mapper.get(cell.getColIndex());
                if (field != null) {
                    Converter converter = converters.get(cell.getColIndex());
                    if (converter == null) {
                        converter = field.getAnnotation(ExcelColumn.class).converter().newInstance();
                        converters.put(cell.getColIndex(), converter);
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
