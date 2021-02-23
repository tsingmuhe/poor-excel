package com.sunchangpeng.poor.excel.utils.convert;

import com.sunchangpeng.poor.excel.cell.StringCellData;

public class StringStringConverter implements Converter<String, StringCellData> {
    @Override
    public boolean support(Class javaTYpe, Class excelType) {
        return String.class == javaTYpe && StringCellData.class == excelType;
    }


    @Override
    public String convert(StringCellData cellData, String defaultValue, ConvertConfig config) {
        if (cellData == null) {
            return defaultValue;
        }

        return cellData.getValue();
    }
}
