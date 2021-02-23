package com.sunchangpeng.poor.excel.utils.convert;

import com.sunchangpeng.poor.excel.cell.DateCellData;

import java.util.Date;

public class DateDateConverter implements Converter<Date, DateCellData> {
    @Override
    public boolean support(Class javaType, Class excelType) {
        return Date.class == javaType && DateCellData.class == excelType;
    }

    @Override
    public Date convert(DateCellData cellData, Date defaultValue, ConvertConfig config) {
        if (cellData == null) {
            return defaultValue;
        }

        return cellData.getValue();
    }
}
