package com.sunchangpeng.poor.excel.write;

import lombok.Getter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.concurrent.atomic.AtomicInteger;

public class WriteSheetHolder {
    @Getter
    private final int sheetIndex;
    @Getter
    private final Sheet sheet;
    @Getter
    private final AtomicInteger currentRowNum = new AtomicInteger(0);

    public WriteSheetHolder(Workbook workbook, int sheetIndex) {
        this.sheetIndex = sheetIndex;
        this.sheet = workbook.createSheet();
    }
}
