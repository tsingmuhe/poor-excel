package com.sunchangpeng.poor.excel.utils.convert;

import com.sunchangpeng.poor.excel.cell.StringCellData;
import org.apache.commons.lang3.StringUtils;

import java.math.BigDecimal;

public class BigDecimalStringConverter implements Converter<BigDecimal, StringCellData> {
    @Override
    public boolean support(Class javaType, Class excelType) {
        return BigDecimal.class == javaType && StringCellData.class == excelType;
    }

    @Override
    public BigDecimal convert(StringCellData cellData, BigDecimal defaultValue, ConvertConfig config) {
        if (cellData == null || StringUtils.isBlank(cellData.getValue())) {
            return defaultValue;
        }

        try {
            return new BigDecimal(cellData.getValue());
        } catch (NumberFormatException e) {
            //pass
            return defaultValue;
        }
    }
}
