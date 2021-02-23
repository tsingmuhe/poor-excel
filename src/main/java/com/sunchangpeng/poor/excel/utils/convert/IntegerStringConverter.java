package com.sunchangpeng.poor.excel.utils.convert;

import com.sunchangpeng.poor.excel.cell.StringCellData;
import org.apache.commons.lang3.StringUtils;

public class IntegerStringConverter implements Converter<Integer, StringCellData> {
    @Override
    public boolean support(Class javaType, Class excelType) {
        return Integer.class == javaType && StringCellData.class == excelType;
    }

    @Override
    public Integer convert(StringCellData cellData, Integer defaultValue, ConvertConfig config) {
        if (cellData == null || StringUtils.isBlank(cellData.getValue())) {
            return defaultValue;
        }

        try {
            return Integer.valueOf(cellData.getValue());
        } catch (NumberFormatException e) {
            //pass?
            return defaultValue;
        }
    }
}
