package com.sunchangpeng.poor.excel.utils.convert;

import com.sunchangpeng.poor.excel.cell.NumberCellData;

import java.util.Date;

import static org.apache.poi.ss.usermodel.DateUtil.getJavaDate;

public class DateNumberConverter implements Converter<Date, NumberCellData> {
    @Override
    public boolean support(Class javaType, Class excelType) {
        return Date.class == javaType && NumberCellData.class == excelType;
    }

    @Override
    public Date convert(NumberCellData cellData, Date defaultValue, ConvertConfig config) {
        if (cellData == null) {
            return defaultValue;
        }

        return getJavaDate(cellData.getValue().doubleValue(), config == null ? false : config.isUse1904windowing());
    }
}
