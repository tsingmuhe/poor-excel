package com.sunchangpeng.poor.excel.utils.convert;

import com.sunchangpeng.poor.excel.cell.NumberCellData;

import java.math.BigDecimal;

public class BigDecimalNumberConverter implements Converter<BigDecimal, NumberCellData> {
    @Override
    public boolean support(Class javaType, Class excelType) {
        return BigDecimal.class == javaType && NumberCellData.class == excelType;
    }

    @Override
    public BigDecimal convert(NumberCellData cellData, BigDecimal defaultValue, ConvertConfig config) {
        if (cellData == null) {
            return defaultValue;
        }

        return cellData.getValue();
    }
}
