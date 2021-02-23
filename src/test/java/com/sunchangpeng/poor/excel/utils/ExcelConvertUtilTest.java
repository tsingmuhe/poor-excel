package com.sunchangpeng.poor.excel.utils;

import com.sunchangpeng.poor.excel.cell.DateCellData;
import com.sunchangpeng.poor.excel.cell.NumberCellData;
import org.junit.Test;

import java.math.BigDecimal;
import java.util.Date;

public class ExcelConvertUtilTest {
    @Test
    public void test() {
        NumberCellData numberCellData = new NumberCellData();
        numberCellData.setValue(new BigDecimal("12.12"));
        BigDecimal result = ExcelConvertUtil.getBigDecimal(numberCellData, null);
        System.out.println(result);

        DateCellData dateCellData = new DateCellData();
        dateCellData.setValue(new Date());
        String result1 = ExcelConvertUtil.getString(dateCellData, null);
        System.out.println(result1);
    }
}