package com.sunchangpeng.poor.excel.utils.convert;

import com.sunchangpeng.poor.excel.cell.StringCellData;
import org.apache.commons.lang3.StringUtils;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class DateStringConverter implements Converter<Date, StringCellData> {
    @Override
    public boolean support(Class javaType, Class excelType) {
        return Date.class == javaType && StringCellData.class == excelType;
    }

    @Override
    public Date convert(StringCellData cellData, Date defaultValue, ConvertConfig config) {
        if (cellData == null || StringUtils.isBlank(cellData.getValue())) {
            return defaultValue;
        }

        String datePattern = DEFAULT_DATE_PATTERN;
        if (config != null && StringUtils.isNotBlank(config.getDatePattern())) {
            datePattern = config.getDatePattern();
        }

        try {
            return new SimpleDateFormat(datePattern).parse(cellData.getValue());
        } catch (ParseException e) {
            //pass
            return defaultValue;
        }
    }
}
