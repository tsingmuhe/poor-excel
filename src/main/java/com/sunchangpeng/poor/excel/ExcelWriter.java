package com.sunchangpeng.poor.excel;

import com.sunchangpeng.poor.excel.utils.ExcelFileUtil;
import com.sunchangpeng.poor.excel.write.ExcelWriteExecutor;

import static org.apache.poi.xssf.streaming.SXSSFWorkbook.DEFAULT_WINDOW_SIZE;

public interface ExcelWriter {
    static ExcelWriteExecutor xlsx() {
        return new ExcelWriteExecutor(ExcelFileUtil.createBook(true));
    }

    static ExcelWriteExecutor xls() {
        return new ExcelWriteExecutor(ExcelFileUtil.createBook(false));
    }

    static ExcelWriteExecutor bigExcel() {
        return bigExcel(DEFAULT_WINDOW_SIZE);
    }

    static ExcelWriteExecutor bigExcel(int windowSize) {
        return new ExcelWriteExecutor(ExcelFileUtil.createSXSSFWorkbook(windowSize));
    }
}