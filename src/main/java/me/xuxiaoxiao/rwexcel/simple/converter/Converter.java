package me.xuxiaoxiao.rwexcel.simple.converter;

import javax.annotation.Nonnull;
import javax.annotation.Nullable;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Objects;
import java.util.regex.Pattern;

/**
 * 数值转换器，将【单元格源字符串值】和【类成员域值】进行互转
 * <ul>
 * <li>[2019/9/19 17:29]XXX：初始创建</li>
 * </ul>
 *
 * @author XXX
 */
public class Converter {
    public static final String D_Y = "yyyy";
    public static final String D_YM = "yyyy-MM";
    public static final String D_YMD = "yyyy-MM-dd";
    public static final String D_YMDHMS = "yyyy-MM-dd HH:mm:ss";
    public static final String D_YMDHMSS = "yyyy-MM-dd HH:mm:ss.SSS";

    public static final Pattern P_Y = Pattern.compile("[1-9]\\d{3}");
    public static final Pattern P_YM = Pattern.compile("[1-9]\\d{3}-\\d{1,2}");
    public static final Pattern P_YMD = Pattern.compile("[1-9]\\d{3}-\\d{1,2}-\\d{1,2}");
    public static final Pattern P_YMDHMS = Pattern.compile("[1-9]\\d{3}-\\d{1,2}-\\d{1,2}\\s\\d{1,2}:\\d{1,2}:\\d{1,2}");
    public static final Pattern P_YMDHMSS = Pattern.compile("[1-9]\\d{3}-\\d{1,2}-\\d{1,2}\\s\\d{1,2}:\\d{1,2}:\\d{1,2}.\\d{1,3}");

    /**
     * 转换异常的处理方法
     */
    public enum ExceptionPolicy {
        /**
         * 继续抛出异常
         */
        Throw,
        /**
         * 简单类型使用默认初始值，引用类型使用null
         */
        Default
    }

    /**
     * 存储SimpleDateFormat映射的ThreadLocal
     */
    private static final ThreadLocal<HashMap<String, SimpleDateFormat>> DATE_FORMATS = new ThreadLocal<>();

    /**
     * 将【单元格源字符串值】转换成【类成员域值】
     *
     * @param field 成员域
     * @param str   【单元格源字符串值】
     * @return 【类成员域值】
     */
    @Nullable
    public Object str2obj(@Nonnull Field field, @Nonnull String str) {
        Class<?> type = field.getType();
        if (exceptionPolicy() == ExceptionPolicy.Throw) {
            if (String.class.equals(type)) {
                return str;
            } else if (int.class.equals(type) || Integer.class.equals(type)) {
                return Integer.parseInt(str);
            } else if (double.class.equals(type) || Double.class.equals(type)) {
                return Float.parseFloat(str);
            } else if (float.class.equals(type) || Float.class.equals(type)) {
                return Float.parseFloat(str);
            } else if (long.class.equals(type) || Long.class.equals(type)) {
                return Long.parseLong(str);
            } else if (boolean.class.equals(type) || Boolean.class.equals(type)) {
                return Float.parseFloat(str);
            } else if (Date.class.equals(type)) {
                if (P_Y.matcher(str).matches()) {
                    return dateParse(D_Y, str);
                }
                if (P_YM.matcher(str).matches()) {
                    return dateParse(D_YM, str);
                }
                if (P_YMD.matcher(str).matches()) {
                    return dateParse(D_YMD, str);
                }
                if (P_YMDHMS.matcher(str).matches()) {
                    return dateParse(D_YMDHMS, str);
                }
                if (P_YMDHMSS.matcher(str).matches()) {
                    return dateParse(D_YMDHMSS, str);
                }
            }
            throw new RuntimeException(String.format("未能将值【%s】解析成【%s】类型的值 ", str, type.getSimpleName()));
        } else {
            if (String.class.equals(type)) {
                return str;
            } else if (int.class.equals(type)) {
                try {
                    return Integer.parseInt(str);
                } catch (Exception e) {
                    return 0;
                }
            } else if (double.class.equals(type)) {
                try {
                    return Double.parseDouble(str);
                } catch (Exception e) {
                    return 0;
                }
            } else if (float.class.equals(type)) {
                try {
                    return Float.parseFloat(str);
                } catch (Exception e) {
                    return 0;
                }
            } else if (long.class.equals(type)) {
                try {
                    return Long.parseLong(str);
                } catch (Exception e) {
                    return 0;
                }
            } else if (boolean.class.equals(type)) {
                try {
                    return Boolean.parseBoolean(str);
                } catch (Exception e) {
                    return false;
                }
            } else if (Date.class.equals(type)) {
                try {
                    if (P_Y.matcher(str).matches()) {
                        return dateParse(D_Y, str);
                    }
                    if (P_YM.matcher(str).matches()) {
                        return dateParse(D_YM, str);
                    }
                    if (P_YMD.matcher(str).matches()) {
                        return dateParse(D_YMD, str);
                    }
                    if (P_YMDHMS.matcher(str).matches()) {
                        return dateParse(D_YMDHMS, str);
                    }
                    if (P_YMDHMSS.matcher(str).matches()) {
                        return dateParse(D_YMDHMSS, str);
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
                return null;
            } else {
                return null;
            }
        }
    }

    /**
     * 将【类成员域值】转换成【单元格源字符串值】
     *
     * @param field 成员域
     * @param obj   【类成员域值】
     * @return 【单元格源字符串值】
     */
    @Nullable
    public String obj2str(@Nonnull Field field, @Nullable Object obj) {
        if (obj == null) {
            return null;
        } else if (obj instanceof String) {
            return (String) obj;
        } else if (obj instanceof Date) {
            return dateFormat(D_YMDHMS, (Date) obj);
        } else if (obj instanceof Integer || obj instanceof Long || obj instanceof Float || obj instanceof Double || obj instanceof Boolean) {
            return String.valueOf(obj);
        } else {
            if (exceptionPolicy() == ExceptionPolicy.Default) {
                return null;
            } else {
                throw new RuntimeException(String.format("未能将【%s】类型值转换成字符串", obj.getClass().getSimpleName()));
            }
        }
    }

    /**
     * 转换异常的处理方法
     *
     * @return 转换异常的处理方法
     */
    @Nonnull
    public ExceptionPolicy exceptionPolicy() {
        return ExceptionPolicy.Default;
    }

    /**
     * 将日期字符串转换成相应的date对象，线程安全
     *
     * @param format  格式字符串
     * @param dateStr 日期字符串
     * @return 相应的date对象
     */
    @Nonnull
    public Date dateParse(@Nonnull String format, @Nonnull String dateStr) {
        Objects.requireNonNull(format);
        Objects.requireNonNull(dateStr);
        HashMap<String, SimpleDateFormat> formats = DATE_FORMATS.get();
        if (formats == null) {
            formats = new HashMap<>();
            DATE_FORMATS.set(formats);
        }
        SimpleDateFormat dateFormat = formats.get(format);
        if (dateFormat == null) {
            dateFormat = new SimpleDateFormat(format);
            formats.put(format, dateFormat);
        }
        try {
            return dateFormat.parse(dateStr);
        } catch (Exception e) {
            e.printStackTrace();
            throw new IllegalArgumentException(String.format("日期:%s 不符合格式:%s", dateStr, format));
        }
    }

    /**
     * 将date对象转换成相应格式的字符串，线程安全
     *
     * @param format 格式字符串
     * @param date   date对象
     * @return 相应格式的字符串
     */
    @Nonnull
    public static String dateFormat(@Nonnull String format, @Nonnull Date date) {
        Objects.requireNonNull(format);
        Objects.requireNonNull(date);
        HashMap<String, SimpleDateFormat> formats = DATE_FORMATS.get();
        if (formats == null) {
            formats = new HashMap<>();
            DATE_FORMATS.set(formats);
        }
        SimpleDateFormat dateFormat = formats.get(format);
        if (dateFormat == null) {
            dateFormat = new SimpleDateFormat(format);
            formats.put(format, dateFormat);
        }
        return dateFormat.format(date);
    }
}
