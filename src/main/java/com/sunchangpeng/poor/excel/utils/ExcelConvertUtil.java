package com.sunchangpeng.poor.excel.utils;

import com.sunchangpeng.poor.excel.cell.CellData;
import com.sunchangpeng.poor.excel.utils.convert.*;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelConvertUtil {
    private static final List<Converter> CONVERTER_LIST = new ArrayList<>();

    static {
        CONVERTER_LIST.add(new BigDecimalNumberConverter());
        CONVERTER_LIST.add(new BigDecimalStringConverter());
        CONVERTER_LIST.add(new DateDateConverter());
        CONVERTER_LIST.add(new DateNumberConverter());
        CONVERTER_LIST.add(new DateStringConverter());
        CONVERTER_LIST.add(new IntegerNumberConverter());
        CONVERTER_LIST.add(new IntegerStringConverter());
        CONVERTER_LIST.add(new LongNumberConverter());
        CONVERTER_LIST.add(new LongStringConverter());
        CONVERTER_LIST.add(new StringDateConverter());
        CONVERTER_LIST.add(new StringNumberConverter());
        CONVERTER_LIST.add(new StringStringConverter());
    }

    private static <T> Converter<T, CellData> getConverter(Class<T> javaType, Class<? extends CellData> cellDataType) {
        return CONVERTER_LIST.stream()
                .filter(item -> item.support(javaType, cellDataType))
                .findFirst().orElse(null);
    }

    public static Integer getInteger(CellData cellData) {
        return getInteger(cellData, null);
    }

    public static Integer getInteger(CellData cellData, Integer defaultValue) {
        Converter<Integer, CellData> converter = getConverter(Integer.class, cellData.getClass());
        return converter == null ? null : converter.convert(cellData, defaultValue);
    }

    public static Long getLong(CellData cellData) {
        return getLong(cellData, null);
    }

    public static Long getLong(CellData cellData, Long defaultValue) {
        Converter<Long, CellData> converter = getConverter(Long.class, cellData.getClass());
        return converter == null ? null : converter.convert(cellData, defaultValue);
    }

    public static BigDecimal getBigDecimal(CellData cellData) {
        return getBigDecimal(cellData, null);
    }

    public static BigDecimal getBigDecimal(CellData cellData, BigDecimal defaultValue) {
        Converter<BigDecimal, CellData> converter = getConverter(BigDecimal.class, cellData.getClass());
        return converter == null ? null : converter.convert(cellData, defaultValue);
    }

    public static Date getDate(CellData cellData, ConvertConfig config) {
        return getDate(cellData, null, config);
    }

    public static Date getDate(CellData cellData, Date defaultValue, ConvertConfig config) {
        Converter<Date, CellData> converter = getConverter(Date.class, cellData.getClass());
        return converter == null ? null : converter.convert(cellData, defaultValue, config);
    }

    public static String getString(CellData cellData) {
        return getString(cellData, null, null);
    }

    public static String getString(CellData cellData, ConvertConfig config) {
        return getString(cellData, null, config);
    }

    public static String getString(CellData cellData, String defaultValue, ConvertConfig config) {
        Converter<String, CellData> converter = getConverter(String.class, cellData.getClass());
        return converter == null ? null : converter.convert(cellData, defaultValue, config);
    }
}
