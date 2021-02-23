package com.sunchangpeng.poor.excel.utils.convert;

import com.sunchangpeng.poor.excel.cell.DateCellData;
import org.apache.commons.lang3.StringUtils;

import java.text.SimpleDateFormat;

public class StringDateConverter implements Converter<String, DateCellData> {
    @Override
    public boolean support(Class javaType, Class excelType) {
        return String.class == javaType && DateCellData.class == excelType;
    }

    @Override
    public String convert(DateCellData cellData, String defaultValue, ConvertConfig config) {
        if (cellData == null || cellData.getValue() == null) {
            return defaultValue;
        }

        String datePattern = DEFAULT_DATE_PATTERN;
        if (config != null && StringUtils.isNotBlank(config.getDatePattern())) {
            datePattern = config.getDatePattern();
        }

        try {
            return new SimpleDateFormat(datePattern).format(cellData.getValue());
        } catch (IllegalArgumentException e) {
            //pass?
            return defaultValue;
        }
    }
}
