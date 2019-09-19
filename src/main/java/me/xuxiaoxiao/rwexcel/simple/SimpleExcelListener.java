package me.xuxiaoxiao.rwexcel.simple;

import me.xuxiaoxiao.rwexcel.ExcelCell;
import me.xuxiaoxiao.rwexcel.ExcelRow;
import me.xuxiaoxiao.rwexcel.ExcelSheet;
import me.xuxiaoxiao.rwexcel.reader.ExcelReader;
import me.xuxiaoxiao.rwexcel.simple.converter.Converter;
import me.xuxiaoxiao.rwexcel.simple.converter.StrConverter;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;
import java.util.TreeMap;

/**
 * 请填写类的描述
 * <ul>
 * <li>[2019/9/18 13:13]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public abstract class SimpleExcelListener<T> implements ExcelReader.Listener {
    private final int cache;
    private final List<T> list;
    private final Class<T> clazz;
    private final boolean adaptive;

    protected final TreeMap<Integer, Field> mapper = new TreeMap<>();

    private ExcelSheet sheet = null;
    private int rowStart = -1, rowEnd = -1;
    private List<ExcelCell> cells = new LinkedList<>();

    public SimpleExcelListener() {
        this(100);
    }

    public SimpleExcelListener(int cache) {
        this.cache = cache;
        this.list = new ArrayList<>(cache);
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
                        if (mapper.containsKey(excelColumn.index())) {
                            throw new IllegalArgumentException(String.format("ExcelColumn列冲突，有多个序号为 %d 的列", excelColumn.index()));
                        } else {
                            mapper.put(excelColumn.index(), field);
                        }
                    } else if (!adaptive) {
                        throw new IllegalArgumentException(String.format("ExcelColumn的列 %s 未设置序号", excelColumn.title()));
                    }
                }
            }
        }
        this.adaptive = adaptive;
    }

    @Override
    public final void onSheet(@Nonnull ExcelSheet sheet) {
        this.sheet = sheet;
    }

    @Override
    public final void onRow(@Nonnull ExcelRow row) {
        try {
            int shtIndex = sheet.getShtIndex();
            while (rowEnd < row.getRowIndex() - 1) {
                //处理空白行
                if (!rowSkip(shtIndex, rowEnd, null)) {
                    list.add(rowEntity(shtIndex, rowEnd, null, null));
                    if (list.size() >= cache) {
                        flush();
                    }
                }
                rowEnd++;
            }
            if (!rowSkip(shtIndex, row.getRowIndex(), row)) {
                list.add(rowEntity(shtIndex, row.getRowIndex(), row, cells));
                if (list.size() >= cache) {
                    flush();
                }
            }
            cells.clear();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Override
    public final void onCell(@Nonnull ExcelCell cell) {
        cells.add(cell);
    }

    private void flush() {
        onList(sheet.getShtIndex(), rowStart, rowEnd, list);
        rowStart = rowEnd;
        list.clear();
    }

    protected Class<T> entityClass() {
        if (this.clazz == null) {
            Class<?> listenerClass = this.getClass();
            if (listenerClass.getGenericSuperclass() instanceof ParameterizedType) {
                Type[] typeArguments = ((ParameterizedType) listenerClass.getGenericSuperclass()).getActualTypeArguments();
                if (typeArguments != null && typeArguments.length > 0) {
                    java.lang.reflect.Type paramType = typeArguments[0];
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

    @Nullable
    protected Field entityField(int rowIndex, int colIndex) {
        return mapper.get(colIndex);
    }

    protected boolean rowSkip(int shtIndex, int rowIndex, @Nullable ExcelRow row) {
        if (rowIndex == 1 || adaptive) {
            //自动寻列模式下，要通过第一行（标题行）找一下属性和列的对应关系
        }
        return rowIndex == 0 || row == null;
    }

    protected T rowEntity(int shtIndex, int rowIndex, @Nullable ExcelRow row, @Nullable List<ExcelCell> cells) throws Exception {
        if (row == null || cells == null || cells.isEmpty()) {
            return null;
        } else {
            T entity = this.clazz.newInstance();
            for (ExcelCell cell : cells) {
                Field field = entityField(cell.getRowIndex(), cell.getColIndex());
                if (field != null) {
                    Class<? extends Converter<?>> cClass = field.getAnnotation(ExcelColumn.class).converter();
                    if (StrConverter.class.equals(cClass)) {
                        field.set(entity, cell.getStrValue());
                    } else {
                        field.set(entity, cClass.newInstance().convert(cell.getStrValue()));
                    }
                }
            }
            return entity;

        }
    }

    protected abstract void onList(int shtIndex, int rowStart, int rowEnd, List<T> list);
}