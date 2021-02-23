package com.sunchangpeng.poor.excel.utils.convert;

import com.sunchangpeng.poor.excel.cell.StringCellData;
import org.apache.commons.lang3.StringUtils;

public class LongStringConverter implements Converter<Long, StringCellData> {
    @Override
    public boolean support(Class javaType, Class excelType) {
        return Long.class == javaType && StringCellData.class == excelType;
    }

    @Override
    public Long convert(StringCellData cellData, Long defaultValue, ConvertConfig config) {
        if (cellData == null || StringUtils.isBlank(cellData.getValue())) {
            return defaultValue;
        }

        try {
            return Long.valueOf(cellData.getValue());
        } catch (NumberFormatException e) {
            //pass?
            return defaultValue;
        }
    }
}
