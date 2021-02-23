package com.sunchangpeng.poor.excel.utils.convert;

public interface Converter<T, R> {
    public static String DEFAULT_DATE_PATTERN = "yyyy-MM-dd HH:mm:ss";

    public boolean support(Class javaType, Class excelType);

    default public T convert(R cellData) {
        return convert(cellData, null, null);
    }

    default public T convert(R cellData, T defaultValue) {
        return convert(cellData, defaultValue, null);
    }

    default public T convert(R cellData, ConvertConfig config) {
        return convert(cellData, null, config);
    }

    public T convert(R cellData, T defaultValue, ConvertConfig config);
}
