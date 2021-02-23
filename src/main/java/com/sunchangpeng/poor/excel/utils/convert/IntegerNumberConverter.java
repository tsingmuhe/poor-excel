package com.sunchangpeng.poor.excel.utils.convert;

import com.sunchangpeng.poor.excel.cell.NumberCellData;

public class IntegerNumberConverter implements Converter<Integer, NumberCellData> {
    @Override
    public boolean support(Class javaType, Class excelType) {
        return Integer.class == javaType && NumberCellData.class == excelType;
    }

    @Override
    public Integer convert(NumberCellData cellData, Integer defaultValue, ConvertConfig config) {
        if (cellData == null || cellData.getValue() == null) {
            return defaultValue;
        }

        return cellData.getValue().intValue();
    }
}
