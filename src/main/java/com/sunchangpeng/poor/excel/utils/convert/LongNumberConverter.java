package com.sunchangpeng.poor.excel.utils.convert;

import com.sunchangpeng.poor.excel.cell.NumberCellData;

public class LongNumberConverter implements Converter<Long, NumberCellData> {
    @Override
    public boolean support(Class javaType, Class excelType) {
        return Long.class == javaType && NumberCellData.class == excelType;
    }

    @Override
    public Long convert(NumberCellData cellData, Long defaultValue, ConvertConfig config) {
        if (cellData == null || cellData.getValue() == null) {
            return defaultValue;
        }

        return cellData.getValue().longValue();
    }
}
