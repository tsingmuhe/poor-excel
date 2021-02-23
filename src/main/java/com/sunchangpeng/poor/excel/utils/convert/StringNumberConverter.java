package com.sunchangpeng.poor.excel.utils.convert;

import com.sunchangpeng.poor.excel.cell.NumberCellData;

public class StringNumberConverter implements Converter<String, NumberCellData> {
    @Override
    public boolean support(Class javaType, Class excelType) {
        return String.class == javaType && NumberCellData.class == excelType;
    }

    @Override
    public String convert(NumberCellData cellData, String defaultValue, ConvertConfig config) {
        if (cellData == null || cellData.getValue() == null) {
            return defaultValue;
        }

        return cellData.getValue().stripTrailingZeros().toPlainString();
    }
}
